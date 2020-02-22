import click
import SRUMParse
from openpyxl import Workbook

@click.command()
@click.option("--input", "-i", required=True, type=click.Path(exists=True,
                                                              dir_okay=False,
                                                              writable=False))

@click.option("--output", "-o", required=True, type=click.Path(exists=True,
                                                               file_okay=False))

@click.option("--include-registry", "-r", required=False, type=click.Path(exists=True,
                                                                          file_okay=False))

@click.option("--force-overwrite", "-f", is_flag=True, default=False)

@click.option("--omit-processed", "-p", is_flag=True, default=False)

@click.option("--only-processed", "-P", is_flag=True, default=False)

@click.pass_context
def export_xlsx(ctx, input_path, output_path, include_registry, force_overwrite, omit_processed, only_processed):
    "Parse and export SRUM data from the --input provided ESE database file " \
    "to an .xlsx file in the --output directory, optionally including the SRUM " \
    "entries found in the registry in the folder provided by --include-registry."

    sourceFile = open(input_path, "r")

    if include_registry is None:
        parser = SRUMParse.SRUMParser(sourceFile)
    else:
        parser = SRUMParse.SRUMParser(sourceFile, registryFolder=include_registry)


    out_workbook = Workbook()

    if not only_processed:
        for table in parser.raw_tables:
            worksheet = out_workbook.create_sheet(table.name)

            works

    for row in parser and

    

    print(input, output, force_overwrite)


export_xlsx()