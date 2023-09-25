#region Namespaces
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Reflection;

#endregion

namespace Int_Mod1_Sched
{
    [Transaction(TransactionMode.Manual)]
    public class Command1 : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            // this is a variable for the Revit application
            UIApplication uiapp = commandData.Application;

            // this is a variable for the current Revit model
            Document doc = uiapp.ActiveUIDocument.Document;

            using (Transaction t = new Transaction(doc))
            {
                t.Start("Starting Mod 1 Challenge");

                //Gather Room Data
                FilteredElementCollector roomCollector = new FilteredElementCollector(doc);
                roomCollector.OfCategory(BuiltInCategory.OST_Rooms);
                roomCollector.WhereElementIsNotElementType();
                IList<ElementId> roomIds = roomCollector.ToElementIds() as IList<ElementId>;

                //Get Room Department Parameter
                List<string> deptList = new List<string>();
                foreach (ElementId roomId in roomIds)
                {
                    Room room = doc.GetElement(roomId) as Room;
                    string roomDept = GetParameterValueAsString(room, "Department");
                    deptList.Add(roomDept.ToString());
                }
                List<string> uniqueDepts = deptList.Distinct().ToList();

                //Get Room Parameters
                Element roomInstance = roomCollector.FirstElement();
                Parameter roomNumParam = roomInstance.LookupParameter("Number");
                Parameter roomNameParam = roomInstance.LookupParameter("Name");
                Parameter roomDeptParam = roomInstance.LookupParameter("Department");
                Parameter roomCommentParam = roomInstance.LookupParameter("Comments");
                Parameter roomAreaParam = roomInstance.get_Parameter(BuiltInParameter.ROOM_AREA);
                Parameter roomLevelParam = roomInstance.LookupParameter("Level");

                int i = 0;
                ViewSchedule[] schedules;
                schedules = new ViewSchedule[uniqueDepts.Count];
                foreach (string depts in uniqueDepts)
                {
                    //Create Schedules
                    ElementId catId = new ElementId(BuiltInCategory.OST_Rooms);
                    schedules[i] = ViewSchedule.CreateSchedule(doc, catId);
                    schedules[i].Name = "Dept - " + depts;

                    //Create Schedule Fields
                    ScheduleField roomNumField = schedules[i].Definition.AddField(ScheduleFieldType.Instance, roomNumParam.Id);
                    ScheduleField roomNameField = schedules[i].Definition.AddField(ScheduleFieldType.Instance, roomNameParam.Id);
                    ScheduleField roomDeptField = schedules[i].Definition.AddField(ScheduleFieldType.Instance, roomDeptParam.Id);
                    ScheduleField roomCommentField = schedules[i].Definition.AddField(ScheduleFieldType.Instance, roomCommentParam.Id);
                    ScheduleField roomAreaField = schedules[i].Definition.AddField(ScheduleFieldType.ViewBased, roomAreaParam.Id);
                    roomAreaField.DisplayType = ScheduleFieldDisplayType.Totals;
                    ScheduleField roomLevelField = schedules[i].Definition.AddField(ScheduleFieldType.Instance, roomLevelParam.Id);
                    roomLevelField.IsHidden = true;

                    //Create Schedule Filter
                    ScheduleFilter deptFilter = new ScheduleFilter(roomDeptField.FieldId, ScheduleFilterType.Equal, depts.ToString());
                    schedules[i].Definition.AddFilter(deptFilter);
                    
                    //Group Schedule by Level
                    ScheduleSortGroupField levelSort = new ScheduleSortGroupField(roomLevelField.FieldId);
                    levelSort.ShowHeader = true;
                    levelSort.ShowFooter = true;
                    levelSort.ShowBlankLine = true;
                    schedules[i].Definition.AddSortGroupField(levelSort);

                    //Sort by Room Name
                    ScheduleSortGroupField nameSort = new ScheduleSortGroupField(roomNameField.FieldId);
                    schedules[i].Definition.AddSortGroupField(nameSort);

                    //Set Totals
                    schedules[i].Definition.ShowGrandTotal = true;
                    schedules[i].Definition.ShowGrandTotalTitle = true;
                    schedules[i].Definition.ShowGrandTotalCount = true;
                    i++;
                }
                //BONUS - Create "All Departments" Schedule
                ElementId roomCatId = new ElementId(BuiltInCategory.OST_Rooms);
                ViewSchedule allDeptSchedule = ViewSchedule.CreateSchedule(doc, roomCatId);
                allDeptSchedule.Name = "All Departments";

                //Create Schedule Fields
                ScheduleField allDeptField = allDeptSchedule.Definition.AddField(ScheduleFieldType.Instance, roomDeptParam.Id);
                ScheduleField allDeptAreaField = allDeptSchedule.Definition.AddField(ScheduleFieldType.ViewBased, roomAreaParam.Id);
                allDeptAreaField.DisplayType = ScheduleFieldDisplayType.Totals;

                //Sort
                ScheduleSortGroupField deptSort = new ScheduleSortGroupField(allDeptField.FieldId);
                allDeptSchedule.Definition.AddSortGroupField(deptSort);

                //Set Totals
                allDeptSchedule.Definition.IsItemized = false;
                allDeptSchedule.Definition.ShowGrandTotal = true;
                allDeptSchedule.Definition.ShowGrandTotalTitle = true;

                t.Commit();
            }

            return Result.Succeeded;
        }
        internal static string GetParameterValueAsString(Element element, string paramName)
        {
            IList<Parameter> paramList = element.GetParameters(paramName);
            Parameter myParam = paramList.First();

            return myParam.AsString();
        }
        internal static PushButtonData GetButtonData()
        {
            // use this method to define the properties for this command in the Revit ribbon
            string buttonInternalName = "btnCommand1";
            string buttonTitle = "Button 1";

            ButtonDataClass myButtonData1 = new ButtonDataClass(
                buttonInternalName,
                buttonTitle,
                MethodBase.GetCurrentMethod().DeclaringType?.FullName,
                Properties.Resources.Blue_32,
                Properties.Resources.Blue_16,
                "This is a tooltip for Button 1");

            return myButtonData1.Data;
        }
    }
}
