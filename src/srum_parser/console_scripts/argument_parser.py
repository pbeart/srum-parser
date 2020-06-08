from argparse import ArgumentParser, Namespace


def default() -> Namespace:
    parser = ArgumentParser()
    parser.add_argument("files", nargs='*', dest="files", default=[])
    parser.add_argument("--only-processed",
                        "-P",
                        dest="only_processed",
                        default=False,
                        action="store_true",
                        help="")
    parser.add_argument("--omit-processed",
                        "-p",
                        dest="omit_processed",
                        default=False,
                        action="store_true",
                        help="")
    parser.add_argument(
        "-o",
        "--output-dir",
        dest="output_dir",
        help=
        "The directory to save the excel files to (defaults to current dir).",
        default=".")
    return parser.parse_args()
