from cover_letters_processor import CoverLetterProcessor
processor = CoverLetterProcessor(
    folder_path="cover_letters",
    excel_file="so_yeu_ly_lich.xlsx"
)
processor.initialize_excel()
processor.wb.save(processor.excel_file)
processor.process_documents()