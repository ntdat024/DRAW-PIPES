# -*- coding: utf-8 -*-
#region library
import clr 
import os
import sys

clr.AddReference("System")
clr.AddReference("System.Data")
clr.AddReference("RevitServices")
clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")
clr.AddReference('PresentationCore')
clr.AddReference('PresentationFramework')
clr.AddReference("System.Windows.Forms")

import math
import System
import RevitServices
import Autodesk
import Autodesk.Revit
import Autodesk.Revit.DB

from Autodesk.Revit.UI import *
from Autodesk.Revit.DB import *
from Autodesk.Revit.UI.Selection import *
from Autodesk.Revit.DB.Mechanical import *
from Autodesk.Revit.DB.Plumbing import *

from System.Collections.Generic import *


#endregion

#region revit infor
# Get the directory path of the script.py & the Window.xaml
dir_path = os.path.dirname(os.path.realpath(__file__))
#xaml_file_path = os.path.join(dir_path, "Window.xaml")

#Get UIDocument, Document, UIApplication, Application
uidoc = __revit__.ActiveUIDocument
uiapp = UIApplication(uidoc.Document.Application)
app = uiapp.Application
doc = uidoc.Document
activeView = doc.ActiveView
#endregion

class FilterPipe(ISelectionFilter):
    def AllowElement(self, element):
        if element.Category.Name == "Pipes":
             return True
        else:
             return False
       
    def AllowReference(self, reference, position):
        return True

#region public method
class Utils:
    def __init__(self):
        pass

    def is_point_inside (self, list_pipes, pick_point):
        max = self.get_max_distance(list_pipes, pick_point)
        list_xyz = self.get_list_point(list_pipes, pick_point)

        point1 = XYZ()
        point2 = XYZ()
        for i in range(len(list_xyz)):
            p1 = list_xyz[i]
            for p2 in list_xyz:
                value = max - p2.DistanceTo(p1)
                if round(value, 0) == 0:
                    point1 = p1
                    point2 = p2
                    break

        pZ = XYZ(pick_point.X, pick_point.Y, point1.Z)
        distance = pZ.DistanceTo(point1) + pZ.DistanceTo(point2)
        
        if round(distance - max , 0) == 0: return True
        return False
    

    def is_none_slope (self, pipe):
        slope = pipe.get_Parameter(BuiltInParameter.RBS_PIPE_SLOPE).AsDouble()
        if slope == 0: return True
        return False

    def create_extend_line (self, line, distance_to_extend):
        sp = line.GetEndPoint(0)
        ep = line.GetEndPoint(1)
        normalize = (ep - sp).Normalize()
        nsp = sp - normalize * distance_to_extend
        nep = ep + normalize * distance_to_extend
        return Line.CreateBound(nsp, nep)
    
    def line_intersect_plane (self, line, plane):
        normal = plane.Normal
        origin = plane.Origin

        sp = line.GetEndPoint(0)
        ep = line.GetEndPoint(1)
        normalize = (ep - sp).Normalize()
        distance = (normal.DotProduct(origin) - normal.DotProduct(sp)) / normal.DotProduct(normalize)
        point = sp + distance * normalize
        return point
    
    def get_list_point (self, list_pipes, pick_point):
        line_0 = list_pipes[0].Location.Curve
        plane = Plane.CreateByNormalAndOrigin(line_0.Direction, pick_point)

        list_xyz = []
        for pipe in list_pipes:
            line = pipe.Location.Curve
            extend_line =  self.create_extend_line(line,500)
            point = self.line_intersect_plane(extend_line, plane)
            list_xyz.append(point)

        return list_xyz
    
    def get_max_distance (self, list_pipes, pick_point):
        list_xyz = self.get_list_point(list_pipes, pick_point)
        list_distance = []
        
        for i in range(len(list_xyz)):
            p1 = list_xyz[i]
            for p2 in list_xyz:
                list_distance.append(p1.DistanceTo(p2))
        
        return max(list_distance)


    def extend_pipe (self, pipe, point_to_extend):
        line = pipe.Location.Curve
        sp = line.GetEndPoint(0)
        ep = line.GetEndPoint(1)

        dis1 = sp.DistanceTo(point_to_extend)
        dis2 = ep.DistanceTo(point_to_extend)

        dir1 = line.Direction
        dir2 = point_to_extend - sp
        radian = dir1.AngleTo(dir2)
        degree = round(radian * 180 / math.pi, 3)

        if degree == 180 :
            if dis1 > dis2 : line = Line.CreateBound(point_to_extend, sp)
            else: line = Line.CreateBound(point_to_extend, ep)
        if degree == 0:
            if dis1 > dis2 : line = Line.CreateBound(sp, point_to_extend)
            else: line = Line.CreateBound(ep, point_to_extend)

        #update pipe curve
        pipe.Location.Curve = line


    def get_nearest_connector (self, pipe, pick_point):
        list_distance = []
        cs = pipe.ConnectorManager.Connectors
        for cn in cs:
            origin = cn.Origin
            distance = origin.DistanceTo(pick_point)
            list_distance.append(distance)

        min_distance = min(list_distance)
        connector = None
        for cn in cs:
            origin = cn.Origin
            distance = origin.DistanceTo(pick_point)
            value = round(distance - min_distance, 3)
            if value == 0: 
                connector = cn
                break

        return connector


    def get_nearest_point_pipe (self, list_pipes, pick_point):
        list_distance = []
        for pipe in list_pipes:
            connector = self.get_nearest_connector(pipe,pick_point)
            pZ = XYZ(pick_point.X, pick_point.Y, connector.Origin.Z)
            distance = pZ.DistanceTo(connector.Origin)
            list_distance.append(distance)
        
        point_pipe = {}
        min_distance = min(list_distance)
        for pipe in list_pipes:
            connector = self.get_nearest_connector(pipe,pick_point)
            pZ = XYZ(pick_point.X, pick_point.Y, connector.Origin.Z)
            distance = pZ.DistanceTo(connector.Origin)
            value = round(distance - min_distance, 3)
            if value == 0:
                point_pipe = {"point": connector.Origin, "pipe": pipe}
                break
        
        return point_pipe
    
    def distance_pipe_to_point (self, pipe, point):
        location = pipe.Location
        line = location.Curve
        extend_line = self.create_extend_line(line, 500)
        return extend_line.Distance(point)
    
    def get_list_distance(self, list_pipes, pick_point):
        point_pipe = self.get_nearest_point_pipe(list_pipes, pick_point)
        nearest_point = point_pipe.get("point")

        list_distance = []
        for pipe in list_pipes:
            distance = self.distance_pipe_to_point(pipe, nearest_point)
            list_distance.append(distance)
        
        list_distance.sort()
        del list_distance[0]
        return list_distance
    
    def is_line_intersect_pipe (self, pipe, line):
        extendLine = self.create_extend_line(line, 500)
        z = extendLine.GetEndPoint(0).Z

        location = pipe.Location
        pipe_line = location.Curve
        sp = pipe_line.GetEndPoint(0)
        ep = pipe_line.GetEndPoint(1)

        nsp = XYZ(sp.X, sp.Y, z)
        nep = XYZ(ep.X, ep.Y, z)
        newLine = Line.CreateBound(nsp, nep)

        result  = extendLine.Intersect(newLine)
        if result == SetComparisonResult.Overlap: return True
        else: return False
    
    def sort_list_pipes (self, list_pipes, pick_point):
        list_distance = self.get_list_distance(list_pipes, pick_point)
        point_pipe = self.get_nearest_point_pipe(list_pipes, pick_point)
        nearest_point = point_pipe.get("point")
        nearest_pipe = point_pipe.get("pipe")

        sorted_list = []
        for value in list_distance:
            for pipe in list_pipes:
                distance = self.distance_pipe_to_point(pipe, nearest_point)
                if round(distance - value, 3) == 0:
                    sorted_list.append(pipe)
                    break
        
        sorted_list.insert(0,nearest_pipe)
        return sorted_list
    

    def get_point_intersect_pipe_line(self, pipe, line):
        pipe_line = pipe.Location.Curve
        pipe_line_extend = self.create_extend_line(pipe_line, 500)
        z = pipe_line_extend.GetEndPoint(0).Z

        extend_line = self.create_extend_line(line, 500)
        sp = extend_line.GetEndPoint(0)
        ep = extend_line.GetEndPoint(1)
        nsp = XYZ(sp.X, sp.Y, z)
        nep = XYZ(ep.X, ep.Y, z)
        newLine = Line.CreateBound(nep, nsp)

        array = clr.Reference[IntersectionResultArray]()
        result = pipe_line_extend.Intersect(newLine, array)
        if result == SetComparisonResult.Overlap: return array.Item[0].XYZPoint
        else: return None
    
    def create_elbow_fitting (self, pipe1, pipe2):
        cS1 = pipe1.ConnectorManager.Connectors
        cS2 = pipe2.ConnectorManager.Connectors
        for cn1 in cS1:
            for cn2 in cS2:
                distance = cn1.Origin.DistanceTo(cn2.Origin)
                if round(distance,3) == 0:
                    doc.Create.NewElbowFitting(cn1, cn2)
                    break
        
    
#endregion

#region action      
class Action: 
    def __init__(self):
        pass

    def extend_multiple_pipes (self, list_pipes, pick_point):
        try:
            t = Transaction(doc, " ")
            t.Start()
            list_xyz = Utils().get_list_point(list_pipes, pick_point)
            for i in range(len(list_pipes)):
                pipe = list_pipes[i]
                point_to_extend = list_xyz[i]
                Utils().extend_pipe(pipe, point_to_extend)
            t.Commit()
        except Exception as e:
            TaskDialog.Show("Extend Multiple Pipes", str(e))
        

    def create_new_pipes (self, list_pipes, pick_point):
        pipe_point = Utils().get_nearest_point_pipe(list_pipes, pick_point)
        nearest_point = pipe_point.get("point")
        nearest_pipe = pipe_point.get("pipe")

        new_point = XYZ(pick_point.X, pick_point.Y, nearest_point.Z)
        original_line = Line.CreateBound(nearest_point, new_point)
        list_distance = Utils().get_list_distance(list_pipes, new_point)
       
        list_line_1 = []
        list_line_2 = []
        for value in list_distance:
            offset_line_1 = original_line.CreateOffset(value, XYZ.BasisZ)
            offset_line_2 = original_line.CreateOffset(-value, XYZ.BasisZ)
            list_line_1.append(offset_line_1)
            list_line_2.append(offset_line_2)
        
        #filter list lines
        list_line_final = []
        isIntersect = Utils().is_line_intersect_pipe(nearest_pipe, list_line_1[0])
        if isIntersect: list_line_final = list_line_2
        else:  list_line_final = list_line_1
        list_line_final.insert(0, original_line)

        #sort list pipes
        list_pipe_sorted = Utils().sort_list_pipes(list_pipes, pick_point)


        #transaction
        list_new_pipes = []
        try:
            t = Transaction(doc, " ")
            t.Start()

            for i in range(len(list_pipe_sorted)):
                pipe_i = list_pipe_sorted[i]
                line_i = list_line_final[i]
                
                #extend pipe
                intersect_point = Utils().get_point_intersect_pipe_line(pipe_i, line_i)
                Utils().extend_pipe(pipe_i, intersect_point)

                #get pipe infor
                systemTypeId = pipe_i.get_Parameter(BuiltInParameter.RBS_PIPING_SYSTEM_TYPE_PARAM).AsElementId()
                levelId = pipe_i.get_Parameter(BuiltInParameter.RBS_START_LEVEL_PARAM).AsElementId()
                diameter = pipe_i.get_Parameter(BuiltInParameter.RBS_PIPE_DIAMETER_PARAM).AsDouble()
                typeId = pipe_i.GetTypeId()

                #create new pipe
                newPipe = Pipe.Create(doc, systemTypeId, typeId, levelId, intersect_point, line_i.GetEndPoint(1))
                newPipe.get_Parameter(BuiltInParameter.RBS_PIPE_DIAMETER_PARAM).Set(diameter)
                list_new_pipes.append(newPipe)
                Utils().create_elbow_fitting(newPipe, pipe_i)
                
            t.Commit()

        except Exception as e:
            TaskDialog.Show("Extend Multiple Pipes", str(e))
        

        return list_new_pipes
        
        
#endregion

#region main task
class Main:
    def __init__(self):
        pass
    def main_task (self):
        try:
            list_pipes = uidoc.Selection.PickElementsByRectangle(FilterPipe(), "Select Pipes")
            selected_pipes = []
            if len(list_pipes) > 0:
                selected_pipes = [pipe for pipe in list_pipes if Utils().is_none_slope(pipe)]

                if len(selected_pipes) > 0:
                    while True:
                        try:
                            pick_point = uidoc.Selection.PickPoint()
                            isPointInside = Utils().is_point_inside(selected_pipes, pick_point)

                            if isPointInside:
                                Action().extend_multiple_pipes(selected_pipes, pick_point)
                            else:
                                selected_pipes = Action().create_new_pipes(selected_pipes, pick_point)
                        except:
                            break
        except:
            pass
    
            
if __name__ == "__main__":
    Main().main_task()
        
#endregion