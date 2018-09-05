import xlrd
import sys

file_name = sys.argv[1] + ".xls"
workbook = xlrd.open_workbook(file_name)
worksheet = workbook.sheet_by_index(0)

total_rows = worksheet.nrows
total_cols = worksheet.ncols

global tasks_by_user
tasks_by_user = {}

global tasks_by_user_c
tasks_by_user_c = {}

global number_skipped
number_skipped = []


def skip_lines():
    for x in range(total_rows):
        name = worksheet.cell(x, 12).value
        if "|" in name:
            number_skipped.append(x)


def tareas_sin_asignar():
    record = set()
    without_complete = 0

    for x in range(total_rows):
        appearances = 0
        appearances_group = 0
        name = worksheet.cell(x, 12).value

        if x in number_skipped:
            continue
        elif name in record:
            continue
        else:
            record.add(name)

        for y in range(total_rows):
            if(y in number_skipped):
                continue
            elif name == worksheet.cell(y, 12).value:
                appearances += 1
            if name == worksheet.cell(y, 12).value or name in worksheet.cell(y, 12).value:
                appearances_group += 1
            y += 1

        print(name.strip(), appearances)
        if name != "Assigned To":
            tasks_by_user[name]=appearances_group
        if name.strip() == '.':
            without_complete += appearances
        x += 1

    assigned_total=(total_rows - 1) - without_complete - len(number_skipped)
    print('')
    print('Asignadas: ' + str(assigned_total))
    print('Sin asignar: ' + str(without_complete))


def tareas_tiempos():
    record=set()
    with_time=0

    for x in range(total_rows):
        appearances=0
        name=worksheet.cell(x, 19).value

        if x in number_skipped:
            continue
        elif name in record:
            continue
        else:
            record.add(name)

        for y in range(total_rows):
            if y in number_skipped:
                continue
            elif name == worksheet.cell(y, 19).value:
                appearances += 1
            y += 1

        if str(name).strip() != '' and str(name) != 'Time Logged Minutes':
            with_time += appearances

        print(name, appearances)
        x += 1

    print('')
    without_time=(total_rows - 1) - with_time - len(number_skipped)
    print("Con tiempo: " + str(with_time))
    print("Sin tiempo: " + str(without_time))


def tareas_completadas():
    record=set()
    current_record_c={}
    for x in range(total_rows):
        appearances=0
        name=worksheet.cell(x, 22).value

        if x in number_skipped:
            continue
        elif name in record:
            continue
        else:
            record.add(name)

        for y in range(total_rows):
            if y in number_skipped:
                continue
            elif name == worksheet.cell(y, 22).value:
                appearances += 1
            y += 1

        print(name, appearances)
        current_record_c[name]=appearances
        x += 1

    print ''

    record=set()

    current_record_o={}
    for x in range(total_rows):
        appearances=0
        name=worksheet.cell(x, 10).value

        if x in number_skipped:
            continue
        elif name in record:
            continue
        else:
            record.add(name)

        for y in range(total_rows):
            if y in number_skipped:
                continue
            elif name == worksheet.cell(y, 10).value:
                appearances += 1
            y += 1

        print(name, appearances)
        current_record_o[name]=appearances
        x += 1

    print ''

    total_completed_on_time=0

    if 1 in current_record_c:
        total_completed_on_time=current_record_c[1]

    total_completed_late=0
    total_not_completed=0
    for key in current_record_o:
        if key != "completed" and key != "Status":
            total_not_completed += current_record_o[key]

    if 0 in current_record_c:
        total_completed_late=current_record_c[0] - total_not_completed

    print("Completadas a tiempo: " + str(total_completed_on_time))
    print("Completadas tarde: " + str(total_completed_late))
    print("Sin completar: " + str(total_not_completed))


def tareas_completadas_integrante():
    record=set()

    for x in range(total_rows):
        appearances=0
        name=worksheet.cell(x, 17).value

        if x in number_skipped:
            continue
        elif name in record:
            continue
        else:
            record.add(name)

        for y in range(total_rows):
            if y in number_skipped:
                continue
            elif name == worksheet.cell(y, 17).value:
                appearances += 1
            y += 1

        print(name, appearances)
        if name != "Completed By Firstname":
            tasks_by_user_c[name]=appearances
        x += 1


def porcentaje_completitud():
    for key in tasks_by_user:
        assigned_tasks=tasks_by_user[key]
        tasks_complete=0
        percentage_completeness=0
        for key2 in tasks_by_user_c:
            if(key.strip() != '' and key2.strip() != '' and key2.strip() in key.strip()):
                tasks_complete=tasks_by_user_c[key2]

        if assigned_tasks != 0:
            percentage_completeness=(
                float(tasks_complete) / float(assigned_tasks)) * 100.0

        if '|' not in key and key.strip() != '.':
            print(key)
            print("Asignadas: " + str(assigned_tasks))
            print("Completadas: " + str(tasks_complete))
            print("Porcentaje de completitud: " + str(percentage_completeness))
            print('')


skip_lines()
print '---------------------------------------------'
print("Numero filas = ", total_rows - 1 - len(number_skipped))
print '---------------------------------------------'
tareas_sin_asignar()
print '---------------------------------------------'
tareas_tiempos()
print '---------------------------------------------'
tareas_completadas()
print '---------------------------------------------'
tareas_completadas_integrante()
print '---------------------------------------------'
porcentaje_completitud()
print '---------------------------------------------'
