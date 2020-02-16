import xlrd
import datetime

loc = ("asd.xlsx")

wb = xlrd.open_workbook(loc)
sheet = wb.sheet_by_index(0)

rows_total = 0
rows_NoItemLabel = 0
failed_NoItemLabel = 0
rows_ItemLabel = 0
failed_ItemLabel = 0
rows_ExtraPacking = 0
failed_ExtraPacking = 0
rows_Kanban = 0
failed_Kanban = 0
rows_Trolley = 0
failed_Trolley = 0
rows_Pallet = 0
failed_Pallet = 0

#for i in range(sheet.nrows):
for i in range(sheet.nrows):
    rows_total += 1
    target = sheet.cell_value(i, 1)
    actual = sheet.cell_value(i, 2)

    target = datetime.datetime.strptime(target, "%d.%m.%Y %H.%M.%S")
    actual = datetime.datetime.strptime(actual, "%d.%m.%Y %H.%M.%S")

    #OTD per order type
    if (sheet.cell_value(i, 5) == "NoItemLabel"):
        rows_NoItemLabel += 1
        if (actual > target):
            #not in time
            failed_NoItemLabel += 1
    if (sheet.cell_value(i, 5) == "ItemLabel"):
        rows_ItemLabel += 1
        if (actual > target):
            #not in time
            failed_ItemLabel += 1
    if (sheet.cell_value(i, 5) == "ExtraPacking"):
        rows_ExtraPacking += 1
        if (actual > target):
            #not in time
            failed_ExtraPacking += 1

    #OTD per delivery unit
    if (sheet.cell_value(i, 6) == "Kanban"):
        rows_Kanban += 1
        if (actual > target):
            #not in time
            failed_Kanban += 1
    if (sheet.cell_value(i, 6) == "Trolley"):
        rows_Trolley += 1
        if (actual > target):
            #not in time
            failed_Trolley += 1
    if (sheet.cell_value(i, 6) == "Pallet"):
        rows_Pallet += 1
        if (actual > target):
            #not in time
            failed_Pallet += 1

OTD_NoItemLabel = (rows_NoItemLabel-failed_NoItemLabel)/rows_NoItemLabel
OTD_ItemLabel = (rows_ItemLabel-failed_ItemLabel)/rows_ItemLabel
OTD_ExtraPacking = (rows_ExtraPacking-failed_ExtraPacking)/rows_ExtraPacking

OTD_Kanban = (rows_Kanban-failed_Kanban)/rows_Kanban
OTD_Trolley = (rows_Trolley-failed_Trolley)/rows_Trolley
OTD_Pallet = (rows_Pallet-failed_Pallet)/rows_Pallet

print("Name Failed Total Percentage")

print("\nNoItemLabel", failed_NoItemLabel, rows_NoItemLabel, round(OTD_NoItemLabel, 5))
print("ItemLabel", failed_ItemLabel, rows_ItemLabel, round(OTD_ItemLabel, 5))
print("ExtraPacking", failed_ExtraPacking, rows_ExtraPacking, round(OTD_ExtraPacking, 5))

print("\nKanban", failed_Kanban, rows_Kanban, round(OTD_Kanban, 5))
print("Trolley", failed_Trolley, rows_Trolley, round(OTD_Trolley, 5))
print("Pallet", failed_Pallet, rows_Pallet, round(OTD_Pallet, 5))

total_ordtyp_otd = (OTD_NoItemLabel+OTD_ItemLabel+OTD_ExtraPacking)/3
total_deluni_otd = (OTD_Kanban+OTD_Trolley+OTD_Pallet)/3

print("\ntotal OTD per order type", round(total_ordtyp_otd, 5))
print("total OTD per delivery unit", round(total_deluni_otd, 5))

total_otd = (total_ordtyp_otd+total_deluni_otd)/2

print("\nTotal regardless of OrderType or DeliveryUnit", round(total_otd, 5))
