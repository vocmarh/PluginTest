using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PluginTest
{
    [TransactionAttribute(TransactionMode.Manual)]
    public class RoomNumberName : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication uiApp = commandData.Application;
            UIDocument uiDoc = uiApp.ActiveUIDocument;
            Document doc = uiDoc.Document;

            var categorySet = new CategorySet();
            categorySet.Insert(Category.GetCategory(doc, BuiltInCategory.OST_Walls));
            categorySet.Insert(Category.GetCategory(doc, BuiltInCategory.OST_StructuralColumns));
            categorySet.Insert(Category.GetCategory(doc, BuiltInCategory.OST_Floors));
            categorySet.Insert(Category.GetCategory(doc, BuiltInCategory.OST_StructuralFoundation));
            categorySet.Insert(Category.GetCategory(doc, BuiltInCategory.OST_StructuralFraming));
            categorySet.Insert(Category.GetCategory(doc, BuiltInCategory.OST_Windows));
            categorySet.Insert(Category.GetCategory(doc, BuiltInCategory.OST_Doors));

            using (Transaction transaction = new Transaction(doc, "Update Room Wall/Floor Comments"))
            {
                transaction.Start();

                CreateShareParameter(uiApp.Application, doc, "ADSK_Группирование", categorySet, BuiltInParameterGroup.PG_IDENTITY_DATA, true);

                UpdateWallFloorComments updater = new UpdateWallFloorComments();
                updater.UpdateWallFloorCommentsWithRoomNames(doc);

                transaction.Commit();
            }

            return Result.Succeeded;
        }
        public class UpdateWallFloorComments
        {
            public void UpdateWallFloorCommentsWithRoomNames(Document doc)
            {

                FilteredElementCollector roomCollector = new FilteredElementCollector(doc, doc.ActiveView.Id).OfCategory(BuiltInCategory.OST_Rooms);

                foreach (Room room in roomCollector)
                {
                    var spatialElementGeometry = new SpatialElementGeometryCalculator(doc).CalculateSpatialElementGeometry(room);
                    var boundarySubFaces = spatialElementGeometry.GetGeometry().Faces.OfType<Face>().SelectMany(x => spatialElementGeometry.GetBoundaryFaceInfo(x));

                    var groupedSubFaces = boundarySubFaces.GroupBy(x => x.SubfaceType).ToDictionary(x => x.Key, y => y.Select(x => x.SpatialBoundaryElement.HostElementId)
                                        .Distinct().Select(x => doc.GetElement(x)));

                    var floors = groupedSubFaces.TryGetValue(SubfaceType.Bottom, out var value) ? value.Select(x => x as Floor).Where(x => x != null) : new List<Floor>();


                    string roomName = room.get_Parameter(BuiltInParameter.ROOM_NAME).AsString();
                    string roomNumber = room.get_Parameter(BuiltInParameter.ROOM_NUMBER).AsString();

                    string roomNameNumber = roomNumber + " " + roomName;

                    IList<IList<BoundarySegment>> segments = room.GetBoundarySegments(new SpatialElementBoundaryOptions());

                    if (segments != null)
                    {
                        foreach (IList<BoundarySegment> segmentList in segments)
                        {
                            foreach (BoundarySegment segment in segmentList)
                            {

                                Element boundingElement = doc.GetElement(segment.ElementId);

                                if (boundingElement != null && boundingElement.Category != null && boundingElement.Category.Name == "Стены")
                                {
                                    Parameter commentParam = boundingElement.LookupParameter("ADSK_Группирование");

                                    if (commentParam != null)
                                    {
                                        string existingComment = commentParam.AsString();

                                        if (!string.IsNullOrEmpty(existingComment))
                                        {
                                            existingComment += ", " + roomNameNumber;
                                            commentParam.Set(existingComment);
                                        }
                                        else
                                        {
                                            commentParam.Set(roomNameNumber);
                                        }
                                    }
                                }
                            }
                        }
                    }

                    foreach (var floor in floors)
                    {
                        Parameter commentParam = floor.LookupParameter("ADSK_Группирование");
                        if (commentParam != null)
                        {
                            string existingComment = commentParam.AsString();
                            if (!string.IsNullOrEmpty(existingComment))
                            {
                                existingComment += ", " + roomNameNumber;
                                commentParam.Set(existingComment);
                            }
                            else
                            {
                                commentParam.Set(roomNameNumber);
                            }
                        }
                    }
                }
            }
        }

        public void CreateShareParameter(Application application, Document doc, string parameterName,
            CategorySet categorySet, BuiltInParameterGroup builtInParameterGroup, bool isInstance)
        {
            DefinitionFile definitionFile = application.OpenSharedParameterFile();

            if (definitionFile == null)
            {
                TaskDialog.Show("Ошибка", "Не найден ФОП");
                return;
            }

            Definition definition = definitionFile.Groups
                .SelectMany(group => group.Definitions)
                .FirstOrDefault(def => def.Name.Equals(parameterName));

            if (definition == null)
            {
                TaskDialog.Show("Ошибка", "Не найден указанный параметр");
                return;
            }

            Binding binding = application.Create.NewTypeBinding(categorySet);
            if (isInstance)
            {
                binding = application.Create.NewInstanceBinding(categorySet);
            }

            BindingMap map = doc.ParameterBindings;
            map.Insert(definition, binding, builtInParameterGroup);
        }
    }
}
