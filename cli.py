import datetime

import click
from openpyxl import Workbook, styles

import SRUMParse
import XLSXOutUtils


ERROR_FONT = styles.Font(color=styles.colors.RED)


@click.command()
@click.option("--input", "-i", required=True, type=click.Path(exists=True,
                                                              dir_okay=False,
                                                              writable=False))

@click.option("--output", "-o", required=True, type=click.Path(exists=False,
                                                               dir_okay=False,
                                                               writable=True))

@click.option("--include-registry", "-r", required=False, type=click.Path(exists=True,
                                                                          file_okay=False))

@click.option("--force-overwrite", "-f", is_flag=True, default=False)

@click.option("--omit-processed", "-p", is_flag=True, default=False)

@click.option("--only-processed", "-P", is_flag=True, default=False)

@click.pass_context
def export_xlsx(ctx, input, output, include_registry, force_overwrite, omit_processed, only_processed):
    "Parse and export SRUM data from the --input provided ESE database file " \
    "to an .xlsx file in the --output directory, optionally including the SRUM " \
    "entries found in the registry in the folder provided by --include-registry."

    source_file = open(input, "rb")

    

    out_workbook = Workbook()

    


    if not only_processed:
        if include_registry is None:
            parser = SRUMParse.SRUMParser(source_file)
        else:
            parser = SRUMParse.SRUMParser(source_file, registryFolder=include_registry)
            
        for table in parser.raw_tables:
            print(table.name)
            sheet_name = SRUMParse.short_table_name(table.name, False)
            worksheet = out_workbook.create_sheet(sheet_name)

            worksheet.append(["Full Table Name:", table.name])

            worksheet.append([col.name for col in table.columns])

            for row in parser.table_rows(table):
                nrow = [XLSXOutUtils.value_to_safe_string(col) for col in row]
                try:
                    worksheet.append(nrow)
                except Exception as e:
                    print("eee D:<")
                    for col in nrow:
                        try:
                            worksheet.append([col])
                        except Exception as e:
                            print(type(col))
                            print([x for x in col])
                            raise e
                    

    if not omit_processed:
        if include_registry is None:
            parser = SRUMParse.SRUMParser(source_file)
        else:
            parser = SRUMParse.SRUMParser(source_file, registryFolder=include_registry)

        SruDbIdMapTable = [table for table in parser.raw_tables if table.name=="SruDbIdMapTable"][0]
        SruDbIdMap = {}
        SruDbIdStartTime = datetime.datetime.now()
        for row in parser.table_rows(SruDbIdMapTable):
            
            value_type = row[0]
            if value_type == 0 or row[2] == None:
                value = None
            elif value_type in [1, 2]: # UTF-16 string
                value = row[2].decode("utf-16")
            elif value_type == 3: # Windows SID
                try:
                    value = SRUMParse.SID_bytes_to_string(row[2])
                except Exception as e:
                    print(row)
                    raise e
            else:
                value = repr(row[2])
            SruDbIdMap[row[1]] = value

        print("Built SruDbIdMapTable in ", datetime.datetime.now()-SruDbIdStartTime)
        # --------------------------------
        # -------- PARSING TABLES --------
        # --------------------------------

        parser = SRUMParse.SRUMParser(source_file)

        for table in parser.raw_tables:
            print(table.name)
            rows = []
            column_names = []

            # --- Network Usage Data Monitor {973F5D5C-1D90-4944-BE8E-24B94231A174} ---
            if table.name == "{973F5D5C-1D90-4944-BE8E-24B94231A174}":
                worksheet = out_workbook.create_sheet(SRUMParse.short_table_name(table.name, True))

                column_inserts = [[3, "EvaluatedAppId"], [5, "EvaluatedUserId"], [6, "InterfaceLuid_IfType"], [6, "InterfaceLuid_NetLuidIndex"]]

                column_names = [col.name for col in table.columns]

                for insert in column_inserts:
                    column_names.insert(insert[0], insert[1])

                worksheet.append(column_names)

                for row, raw_row in zip(parser.table_rows(table), parser.raw_table_rows(table)):
                    out_row = row[:]
                    
                    luid_parsed = SRUMParse.parse_interface_luid(parser.row_element_by_column_name(raw_row, "InterfaceLuid", table))

                    for insert in column_inserts:
                        if insert[1] == "EvaluatedAppId":
                            value = SruDbIdMap[parser.row_element_by_column_name(row, "AppId", table)]
                            #value = parser.row_element_by_column_name(row, "AppId", table)
                        elif insert[1] == "EvaluatedUserId":
                            value = SruDbIdMap[parser.row_element_by_column_name(row, "UserId", table)]
                            #value = parser.row_element_by_column_name(row, "UserId", table)
                        elif insert[1] == "InterfaceLuid_IfType":
                            value = luid_parsed["iftype"]

                        elif insert[1] == "InterfaceLuid_NetLuidIndex":
                            value = luid_parsed["netluid_index"]
                            
                        else:
                            value = "Could not parse value"
                        out_row.insert(insert[0], XLSXOutUtils.value_to_safe_string(value))

                    worksheet.append(out_row)
            else:
                continue # Not one of the handled sheets


            worksheet.append(["(Processed) Full Table Name:", table.name])
            worksheet.append(column_names)
        
        
    out_workbook.remove(out_workbook["Sheet"]) # Remove default sheet
    out_workbook.save(output)


export_xlsx()
