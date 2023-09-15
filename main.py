import fitz
from openpyxl.reader.excel import load_workbook

FILE_TO_READ = (
    "Abstract Book from the 5th World "
    "Psoriasis and Psoriatic Arthritis Conference 2018.pdf"
)
FILE_TO_WRITE = (
    "Data Entry - 5th World Psoriasis & Psoriatic "
    "Arthritis Conference 2018 - Case format (2).xlsx"
)


def scrape(file_path):
    workbook = load_workbook(filename=FILE_TO_WRITE)
    worksheet = workbook.active
    worksheet.delete_rows(7, 10)
    fieldnames = [
        "Name (incl. titles if any mentioned)",
        "Affiliation(s) Name(s)",
        "Person's Location",
        "Session Name",
        "Topic Title",
        "Presentation Abstract"
    ]
    results = []
    pdf = fitz.open(file_path)
    start_page = 43
    end_page = 60
    pages = [p for p in range(start_page, end_page)]
    pdf.select(pages)
    current_page = {k: "" for k in fieldnames}
    for page in pdf:
        blocs_dict = page.get_text("dict", clip=fitz.Rect(50, 50, 545, 730))
        blocks = blocs_dict["blocks"]
        for block in blocks:
            if "lines" in block.keys():
                spans = block["lines"]
                for span in spans:
                    data = span["spans"]
                    for lines in data:
                        if lines["font"] == "TimesNewRomanPS-BoldItal":
                            if current_page["Session Name"]:
                                results.append(current_page)
                                current_page = {k: "" for k in fieldnames}
                            current_page["Session Name"] += lines["text"]
                        elif (
                            lines["font"] == "TimesNewRomanPS-ItalicMT"
                            and lines["size"] == 9.0
                        ):
                            current_page[
                                "Name (incl. titles if any mentioned)"
                            ] += lines["text"]
                        elif (
                            lines["font"] == "TimesNewRomanPS-BoldMT"
                            and (lines["size"] == 9.0
                                 or lines["size"] == 9.899999618530273)
                        ):
                            current_page["Topic Title"] += lines["text"]
                        elif (
                            lines["font"] == "TimesNewRomanPS-ItalicMT"
                            and lines["size"] == 8.0
                        ):
                            current_page["Affiliation(s) Name(s)"] += lines[
                                "text"
                            ]
                        elif lines["font"] == "TimesNewRomanPS-ItalicMT" and (
                            lines["size"] == 5.247000217437744
                            or lines["size"] == 4.664000034332275
                        ):
                            pass
                        else:
                            current_page["Presentation Abstract"] += lines[
                                "text"
                            ]
    pdf.close()
    for result in results:
        values = (result[k] for k in fieldnames)
        worksheet.append(values)
    workbook.save(filename="test_task_beetroot.xlsx")
    return results


scrape(FILE_TO_READ)
