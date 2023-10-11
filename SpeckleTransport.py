from specklepy.api import operations
from specklepy.api.client import SpeckleClient
from specklepy.api.credentials import get_default_account, get_account_from_token
from specklepy.transports.server import ServerTransport
import numpy as np
import pandas as pd
import comtypes.client


def ImportRevitModel(streamID, commitID, client, SapModel):
    # get the specified commit data
    commit = client.commit.get(stream_id=streamID, commit_id=commitID)

    # create an authenticated server transport from the client and receive the commit obj
    transport = ServerTransport(client=client, stream_id=streamID)

    # extract data from model
    res = operations.receive(commit.referencedObject, remote_transport=transport)

    # get the list of levels from the received object
    elements = res["elements"]
    #---------------------------------------------------------------------------------------------------------
    # GRAB REVIT CATEGORIES DATA FROM MODEL
    revit_category = res.elements[3]
    # Create a dictionary to store unique levels by name
    unique_levels = {}
    revit_category = res.elements[3]

    for cat in revit_category.elements:
        level_name = cat.level.name
        if level_name not in unique_levels:
            unique_levels[level_name] = np.round(cat.level.elevation, 0)

    # sort levels by elevation
    sorted_levels = dict(sorted(unique_levels.items(), key= lambda item: item[1]))
    sorted_levels.pop("DATUM") # remove Revit datum from dict
    first_pair = next(iter(sorted_levels.items())) # grab first key value pair

    # create ETABS levels
    base_elev = first_pair[1]
    num_stories = len(sorted_levels)
    story_names = list(sorted_levels.keys())
    # create list of story heights
    story_elevations = list(sorted_levels.values())
    story_heights = [story_elevations[0]] + [story_elevations[i] - story_elevations[i-1] for i in range(1, len(story_elevations))]
    is_master = [False] * num_stories
    sim_story = ["None"] * num_stories
    splice_above = [False] * num_stories
    splice_h = [0] * num_stories

    # define levels in model
    ret = SapModel.Story.SetStories_2(base_elev, num_stories-1, story_names[1:], story_heights[1:], is_master[1:], sim_story[1:], splice_above[1:], splice_h[1:])
    #---------------------------------------------------------------------------------------------------------
    # GRAB FLOOR DATA FROM MODEL
    floors = elements[3].elements

    for index, floor in enumerate(floors):
        if floor.speckle_type == "Objects.BuiltElements.Floor:Objects.BuiltElements.Revit.RevitFloor":
            # grab outline of floor 
            segments = floor.outline.segments
            num_segments = len(segments)

            # intialise outline list
            X = [] 
            Y = []
            Z = []

            for i, segment in enumerate(segments):
                # grab start and end points from current segment
                # remember - start of segment 2 = end of segment 1 
                
                # find if segment is line or arc
                if  segment.speckle_type == "Objects.Geometry.Line":
                    start = segment.start
                else:
                    start = segment.startPoint

                # check if on the final iteration of the segments loop or not
                if i == num_segments:
                    # if on the final iteration, point = start of first segment
                    x = X[0]
                    y = Y[0]
                    z = Z[0]
                else:
                    # all other iterations = grab start point of the segment
                    x = np.round(start.x, 0)
                    y = np.round(start.y, 0)
                    z = np.round(floor.level.elevation, 0)

                X.append(x)
                Y.append(y)
                Z.append(z)

            # use ETABS API to draw current floor    
            ret = SapModel.AreaObj.AddByCoord(num_segments, X, Y, Z, "floor {index}", "Slab1")
        else:
            continue
    #---------------------------------------------------------------------------------------------------------
    # GRAB COLUMN DATA FROM MODEL
    columns = elements[2].elements

    for index, column in enumerate(columns):   
        # grab start and end points of column
        start = column.baseLine.start # grab x, y and z at bottom of column
        end = column.baseLine.end # grab x, y and z at top of column

        # grab column base and top level names
        baseLevel = column.level.name
        topLevel = column.topLevel.name

        # iterate within range of start and end level
        in_range = False
        for index, level in enumerate(story_names):
            if level == baseLevel:
                in_range = True
            if in_range:
                # extract start and end coordinates of column
                startX = np.round(start.x, 0)
                startY = np.round(start.y, 0)
                startZ = story_elevations[index]
                endX = np.round(end.x, 0)
                endY = np.round(end.y, 0)
                endZ = story_elevations[index + 1]
                # add column with ETABS API
                ret = SapModel.FrameObj.AddByCoord(startX, startY, startZ, endX, endY, endZ)
            if story_names[index + 1] == topLevel:
                in_range = False
                break
                
    #---------------------------------------------------------------------------------------------------------
    # GRAB WALL DATA FROM MODEL
    walls = elements[1].elements

    # loop through each wall in the collection
    for i, wall in enumerate(walls):
        if wall.level and wall.topLevel and wall.height >= min(story_heights):
            # grab start and end points of wall
            start = wall.baseLine.start
            end = wall.baseLine.end
            
            # grab column base and top level names
            baseLevel = wall.level.name
            topLevel = wall.topLevel.name

            # iterate within range of start and end level
            in_range = False
            for index, level in enumerate(story_names):
                if level == baseLevel:
                    in_range = True
                if in_range:
                    # extract start and end coordinates of wall
                    startX = np.round(start.x, 0)
                    startY = np.round(start.y, 0)
                    startZ = story_elevations[index]
                    endX = np.round(end.x, 0)
                    endY = np.round(end.y, 0)
                    endZ = story_elevations[index + 1]

                    # DRAW WALL IN ETABS
                    X_coord = [startX, startX, endX, endX]
                    Y_coord = [startY, startY, endY, endY]
                    Z_coord = [startZ, endZ, endZ, startZ]

                    Name = " "
                    ret = SapModel.AreaObj.AddByCoord(4, X_coord, Y_coord, Z_coord, Name, "CW-200-C40")

                if story_names[index + 1] == topLevel:
                    in_range = False
                    break
        else:
            continue

