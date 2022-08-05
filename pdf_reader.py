import PyPDF2

def main():
    pdfFile = open("C:/Users/si-de/Downloads/Bilag-132252-2022.pdf", "rb")
    pdfReader = PyPDF2.PdfReader(pdfFile)
    page = pdfReader._get_page(0)
    pageText = (page.extract_text())

    for line in pageText.splitlines():
        if "Sum debet" in line:
            print(float("".join(line.split(" ")[2:]).replace(",", ".")))

if __name__ == '__main__':
    main()