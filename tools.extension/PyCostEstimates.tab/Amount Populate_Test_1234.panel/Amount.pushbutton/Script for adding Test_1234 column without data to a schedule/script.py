# -*- coding: utf-8 -*-
from pyrevit import revit, forms
import Autodesk.Revit.DB as DB

SHARED_PARAM_NAME = "Test_1234"

collector = DB.FilteredElementCollector(revit.doc).OfClass(DB.ViewSchedule)
schedules = collector.ToElements()

# Find shared parameter element
param_elements = DB.FilteredElementCollector(revit.doc)\
    .OfClass(DB.ParameterElement)\
    .ToElements()

target_param = None
for p in param_elements:
    if p.Name == SHARED_PARAM_NAME:
        target_param = p
        break

if not target_param:
    forms.alert("Shared parameter '{}' not found.".format(SHARED_PARAM_NAME), title="Missing Param")
    raise SystemExit

param_id = target_param.Id
added = []
skipped = []

# Wrap model changes in a Transaction
t = DB.Transaction(revit.doc, "Add '{}' to schedules".format(SHARED_PARAM_NAME))
t.Start()

for schedule in schedules:
    try:
        definition = schedule.Definition
        field_names = [definition.GetField(i).GetName() for i in range(definition.GetFieldCount())]

        if SHARED_PARAM_NAME in field_names:
            skipped.append("{} (already has '{}')".format(schedule.Name, SHARED_PARAM_NAME))
            continue

        definition.AddField(DB.ScheduleFieldType.Instance, param_id)
        added.append(schedule.Name)

    except Exception as e:
        skipped.append("{} (error: {})".format(schedule.Name, str(e)))

t.Commit()

# Summary alert
msg = "✅ Added shared parameter '{}' to:\n".format(SHARED_PARAM_NAME) + "\n".join(added)
if skipped:
    msg += "\n\n⚠️ Skipped:\n" + "\n".join(skipped)
forms.alert(msg, title="Shared Parameter Column Update", warn_icon=False)
