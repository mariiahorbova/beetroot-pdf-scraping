import fitz
from openpyxl.reader.excel import load_workbook


class PdfScraper:
    @staticmethod
    def scrap(input_file, start_page, end_page, columns):
        results = []
        pdf = fitz.open(input_file)
        pages = [p for p in range(start_page, end_page)]
        pdf.select(pages)
        current_page = {k: "" for k in columns}

        for page in pdf:
            blocks_dict = page.get_text(
                "dict",
                clip=fitz.Rect(50, 50, 545, 730)
            )

            blocks = blocks_dict["blocks"]

            for block in blocks:
                if "lines" in block.keys():
                    spans = block["lines"]
                    for span in spans:
                        data = span["spans"]
                        for lines in data:
                            if lines["font"] == "TimesNewRomanPS-BoldItal":
                                if current_page["Session Name"]:
                                    results.append(current_page)
                                    current_page = {k: "" for k in columns}
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
                                and (
                                    lines["size"] == 9.0
                                    or lines["size"] == 9.899999618530273
                                )
                            ):
                                current_page["Topic Title"] += lines["text"]
                            elif (
                                lines["font"] == "TimesNewRomanPS-ItalicMT"
                                and lines["size"] == 8.0
                            ):
                                current_page[
                                    "Affiliation(s) Name(s)"
                                ] += lines["text"]
                            elif (
                                lines["font"] == "TimesNewRomanPS-ItalicMT"
                                and (
                                    lines["size"] == 5.247000217437744
                                    or lines["size"] == 4.664000034332275
                                )
                            ):
                                pass
                            else:
                                current_page[
                                    "Presentation Abstract"
                                ] += lines["text"]
        pdf.close()
        return results

    @staticmethod
    def save_to_xlsx(data, output_file, columns):
        workbook = load_workbook(filename=output_file)
        worksheet = workbook.active
        worksheet.delete_rows(7, 10)
        for result in data:
            values = (result[k] for k in columns)
            worksheet.append(values)
        workbook.save(filename=output_file)


if __name__ == "__main__":
    file_to_read = (
        "Abstract Book from the 5th World "
        "Psoriasis and Psoriatic Arthritis Conference 2018.pdf"
    )
    file_to_write = (
        "Data Entry - 5th World Psoriasis & Psoriatic "
        "Arthritis Conference 2018 - Case format (2).xlsx"
    )
    fieldnames = [
        "Name (incl. titles if any mentioned)",
        "Affiliation(s) Name(s)",
        "Person's Location",
        "Session Name",
        "Topic Title",
        "Presentation Abstract"
    ]
    scraped_data = PdfScraper.scrap(file_to_read, 43, 60, fieldnames)
    PdfScraper.save_to_xlsx(scraped_data, file_to_write, fieldnames)
