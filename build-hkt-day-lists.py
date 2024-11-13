import io
import pathlib
import typing as t

import google.oauth2.credentials
import google.oauth2.service_account
import googleapiclient.errors  # type: ignore
import googleapiclient.http  # type: ignore
import pandas as pd
import pdfkit  # type: ignore
from docx import Document
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls
from docx.shared import Inches
from googleapiclient.discovery import build  # type: ignore

HKT_FILE_NAME = "HomeKitaTage.xlsx"
HKT_FILE_PATH = pathlib.Path(f"/tmp/{HKT_FILE_NAME}")
SCOPES = ["https://www.googleapis.com/auth/drive"]
SERVICE_ACCOUNT_FILE = pathlib.Path("/tmp/elternvertretung-b7713037bac6.json")
if not SERVICE_ACCOUNT_FILE.is_file():
    SERVICE_ACCOUNT_FILE = pathlib.Path(
        "/data/elternvertretung-b7713037bac6.json"
    )
CREDENTIALS = (
    google.oauth2.service_account.Credentials.from_service_account_file(
        SERVICE_ACCOUNT_FILE, scopes=SCOPES
    )
)


def dataframe_to_word(df, docx_file_path):
    document = Document()

    # Set custom margins (e.g., 0.5 inches for top and bottom)
    sections = document.sections
    for section in sections:
        section.top_margin = Inches(0.5)
        section.bottom_margin = Inches(0.5)

    # Add a table with borders
    table = document.add_table(rows=1, cols=len(df.columns))
    table.style = "Table Grid"  # Use a built-in style with borders

    # Add header row
    hdr_cells = table.rows[0].cells
    for i, column in enumerate(df.columns):
        hdr_cells[i].text = str(column)

    # Add data rows
    for index, row in df.iterrows():
        row_cells = table.add_row().cells
        for i, value in enumerate(row):
            row_cells[i].text = str(value)

    # Apply borders to each cell (if needed)
    for row in table.rows:
        for cell in row.cells:
            cell._element.get_or_add_tcPr().append(
                parse_xml(r'<w:shd {} w:fill="FFFFFF"/>'.format(nsdecls("w")))
            )
            cell._element.get_or_add_tcPr().append(
                parse_xml(
                    (
                        r'<w:tcBorders %s><w:top w:val="single" w:sz="4"/>'
                        r'<w:left w:val="single" w:sz="4"/>'
                        r'<w:bottom w:val="single" w:sz="4"/>'
                        r'<w:right w:val="single" w:sz="4"/>'
                        r"</w:tcBorders>"
                    )
                    % nsdecls("w")
                )
            )

    document.save(docx_file_path)


def export_excel_file_from_google_drive(filename) -> None:
    HKT_FILE_PATH.unlink(missing_ok=True)
    file: t.Optional[io.BytesIO] = None
    try:
        service = build("drive", "v3", credentials=CREDENTIALS)
        results = (
            service.files()
            .list(fields="nextPageToken, files(id, name)")
            .execute()
        )
        items = results.get("files", [])

        if not items:
            print("No files found.")
        for item in items:
            if item["name"] != filename:
                continue
            file_id = item["id"]
            request = service.files().get_media(fileId=file_id)
            file = io.BytesIO()
            downloader = googleapiclient.http.MediaIoBaseDownload(
                file, request
            )
            done = False
            while done is False:
                status, done = downloader.next_chunk()
            HKT_FILE_PATH.write_bytes(file.getvalue())
        if not HKT_FILE_PATH.is_file():
            print(f"File {filename} not found.")
    except googleapiclient.errors.HttpError as error:
        print(f"An error occurred: {error}")


def upload_overview_files_to_google_drive(
    daily_overview_file_paths: list[pathlib.Path],
) -> None:
    try:
        service = build("drive", "v3", credentials=CREDENTIALS)
        existing_files = (
            service.files()
            .list(fields="nextPageToken, files(id, name)")
            .execute()
        ).get("files", [])
        for file_path in daily_overview_file_paths:
            if existing_files:
                for existing_file in existing_files:
                    if existing_file["name"] == file_path.stem:
                        service.files().delete(
                            fileId=existing_file["id"]
                        ).execute()
                        existing_files.remove(existing_file)
                        break
            file_metadata = {
                "name": file_path.stem,
                "parents": ["19PV7rVVVA1uPS-LIDU5bdHGYvz3lODXB"],
            }
            media = googleapiclient.http.MediaFileUpload(
                file_path, chunksize=-1
            )
            file = (
                service.files()
                .create(
                    body=file_metadata,
                    media_body=media,
                    fields="id,name,webViewLink",
                )
                .execute()
            )
            print(f"Uploaded {file.get('name')} to {file.get('webViewLink')}.")
            # print(f'Link: {file.get("webViewLink")}')
            # permission = {
            #     "type": "user",
            #     "role": "writer",
            #     "emailAddress": "elternvertretung@bluetezeit-berlin.de"
            # }
            # # https://developers.google.com/drive/api/reference/rest/v3/permissions/create
            # service.permissions().create(
            #     fileId=file.get("id"),
            #     body=permission,
            #     transferOwnership=False,
            #     sendNotificationEmail=False,
            #     supportsAllDrives=True,
            #     # moveToNewOwnersRoot=True,
            # ).execute()
    except googleapiclient.errors.HttpError as error:
        print(f"An error occurred: {error}")


def generate_daily_overview_files() -> list[pathlib.Path]:
    daily_overview_file_paths: list[pathlib.Path] = []
    df = pd.read_excel(HKT_FILE_PATH, sheet_name="HKT Erfassung")
    for group_name, group_df in df.groupby("Group"):
        for no, day in enumerate(
            (
                "Monday",
                "Tuesday",
                "Wednesday",
                "Thursday",
                "Friday",
            ),
            start=1,
        ):
            day_df = group_df[
                (group_df[f"{day}\nmorning"] == 1.0)
                | (group_df[f"{day}\nafternoon"] == 1.0)
            ]
            day_df = day_df.replace(1.0, "Stay at home")
            html_file_path = pathlib.Path(f"/tmp/{group_name}_{no}_{day}.html")
            df = day_df[
                [
                    "Name",
                    "Group",
                    f"{day}\nmorning",
                    f"{day}\nafternoon",
                ]
            ].fillna("")
            df.to_html(html_file_path, index=False)
            pdf_file_path = html_file_path.with_suffix(".pdf")
            options = {"encoding": "UTF-8", "user-style-sheet": "style.css"}
            pdfkit.from_file(
                input=str(html_file_path),
                output_path=str(pdf_file_path),
                options=options,
                verbose=False,
            )
            docx_file_path = html_file_path.with_suffix(".docx")
            dataframe_to_word(df, docx_file_path)
            # html_file_path.unlink(missing_ok=True)
            for file_path in (
                pdf_file_path,
                docx_file_path,
            ):
                if file_path.is_file():
                    daily_overview_file_paths.append(file_path)
    return daily_overview_file_paths


if __name__ == "__main__":
    export_excel_file_from_google_drive(filename=HKT_FILE_NAME)
    if not HKT_FILE_PATH.is_file():
        print(f"File {HKT_FILE_NAME} not found.")
        raise SystemExit(1)
    daily_overview_file_paths = generate_daily_overview_files()
    upload_overview_files_to_google_drive(daily_overview_file_paths)
