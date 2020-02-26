import click
import SRUMParse
from openpyxl import Workbook

import datetime

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

    if include_registry is None:
        parser = SRUMParse.SRUMParser(source_file)
    else:
        parser = SRUMParse.SRUMParser(source_file, registryFolder=include_registry)

    out_workbook = Workbook()

    table_aliases = { # To keep worksheet names under 31 characters
        "{973F5D5C-1D90-4944-BE8E-24B94231A174}": "Network Data Usage Monitor",
        "{D10CA2FE-6FCF-4F6D-848E-B2E99266FA89}": "Application Resource Usage",
        "{DA73FB89-2BEA-4DDC-86B8-6E048C6DA477}": "Energy Estimator ...6DA477}",
        "{DD6636C4-8929-4683-974E-22C046A43763}": "Network Conn. Usage Monitor",
        "{FEE4E14F-02A9-4550-B5CE-5FA2DA202E37}": "Energy Usage",
        "{FEE4E14F-02A9-4550-B5CE-5FA2DA202E37}LT": "Long-term Energy Usage",
        "{D10CA2FE-6FCF-4F6D-848E-B2E99266FA86}": "Push Notifications",
        "{5C8CF1C7-7257-4F13-B223-970EF5939312}": "Energy Estimator ...939312}"
    }


    if not only_processed:

        for table in parser.raw_tables:
            sheet_name = table_aliases[table.name] if table.name in table_aliases else table.name
            print(sheet_name)
            worksheet = out_workbook.create_sheet(sheet_name)

            if sheet_name != table.name:
                full_name = "{0} ({1})".format(table.name, sheet_name)
            else:
                full_name = table.name
            worksheet.append(["Full Table Name:", full_name])

            worksheet.append([col.name for col in table.columns])

            for row in parser.table_rows(table):
                worksheet.append(row)

    if not omit_processed:
        SruDbIdMapTable = [table for table in parser.raw_tables if table.name == "SruDbIdMapTable"][0]
        SruDbIdMap = {}
        for row in parser.table_rows(SruDbIdMapTable):
            value_type = row[0]
            if value_type in [0, 1, 2]: # UTF-16 string
                value = row[2].encode("utf-16")
            elif value_type == 3: # Windows SID
                value = SRUMParse.SID_bytes_to_string(row[2])
            else:
                value = repr(row[2])
            SruDbIdMap[row[1]] = value
        


        network_usage_monitor = 
        print(datetime.datetime.now().strftime("%H:%M:%S.%f"))
        for record in network_usage_monitor.records:
            pass
            
        print("Done!")
        print(datetime.datetime.now().strftime("%H:%M:%S.%f"))
        input()






        
        
    out_workbook.remove(out_workbook["Sheet"])
    out_workbook.save(output)


export_xlsx()
