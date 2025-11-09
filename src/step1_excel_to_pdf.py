"""
Step 1: Excel to PDF conversion module
"""
import xlwings as xw
from pathlib import Path
from typing import List
from logger import logger


def list_all_sheets(excel_path: str) -> List[str]:
    """
    List all sheet names in an Excel file

    Args:
        excel_path: Path to Excel file

    Returns:
        List of sheet names

    Raises:
        FileNotFoundError: If Excel file doesn't exist
        Exception: If Excel operation fails
    """
    excel_file = Path(excel_path)
    if not excel_file.exists():
        raise FileNotFoundError(f"Excel file not found: {excel_path}")

    logger.info(f"Opening Excel file: {excel_path}")

    app = xw.App(visible=False)
    try:
        wb = xw.Book(excel_path)
        sheet_names = [sheet.name for sheet in wb.sheets]
        logger.info(f"Found {len(sheet_names)} sheets: {sheet_names}")
        wb.close()
        return sheet_names
    except Exception as e:
        logger.error(f"Failed to list sheets: {str(e)}")
        raise
    finally:
        app.quit()


def generate_pdfs(
    excel_path: str,
    sheet_names: List[str],
    version: int,
    output_dir: str
) -> List[str]:
    """
    Generate PDF files from Excel sheets

    Args:
        excel_path: Path to Excel file
        sheet_names: List of sheet names to export
        version: Version number
        output_dir: Output directory for PDF files

    Returns:
        List of generated PDF file paths

    Raises:
        FileNotFoundError: If Excel file doesn't exist
        ValueError: If sheet name doesn't exist
        Exception: If PDF generation fails
    """
    excel_file = Path(excel_path)
    if not excel_file.exists():
        raise FileNotFoundError(f"Excel file not found: {excel_path}")

    # Create output directory
    output_path = Path(output_dir)
    output_path.mkdir(parents=True, exist_ok=True)

    logger.info(f"Starting PDF generation for {len(sheet_names)} sheets")

    app = xw.App(visible=False)
    generated_files = []

    try:
        wb = xw.Book(excel_path)
        base_name = excel_file.stem  # Filename without extension

        # Verify all sheet names exist
        available_sheets = [sheet.name for sheet in wb.sheets]
        for sheet_name in sheet_names:
            if sheet_name not in available_sheets:
                raise ValueError(f"Sheet '{sheet_name}' not found in Excel file")

        # Generate PDF for each sheet
        for sheet_name in sheet_names:
            try:
                logger.info(f"Processing sheet: {sheet_name}")
                ws = wb.sheets[sheet_name]

                # Configure page setup for PDF
                ps = ws.api.PageSetup

                # Set page orientation to landscape
                ps.Orientation = xw.constants.PageOrientation.xlLandscape

                # Force all columns to fit in 1 page width
                ps.Zoom = False  # Disable zoom, use FitToPages instead
                ps.FitToPagesWide = 1      # All columns must fit in 1 page width
                ps.FitToPagesTall = False  # False = unlimited pages for height (rows can span multiple pages)

                # Set smaller margins to maximize content area
                ps.LeftMargin = ws.api.Application.InchesToPoints(0.25)
                ps.RightMargin = ws.api.Application.InchesToPoints(0.25)
                ps.TopMargin = ws.api.Application.InchesToPoints(0.25)
                ps.BottomMargin = ws.api.Application.InchesToPoints(0.25)
                ps.HeaderMargin = ws.api.Application.InchesToPoints(0.2)
                ps.FooterMargin = ws.api.Application.InchesToPoints(0.2)

                # Generate PDF filename
                pdf_name = f"{base_name}_{sheet_name}_V{version}.pdf"
                pdf_path = output_path / pdf_name

                # Export to PDF
                logger.info(f"Exporting to: {pdf_path}")
                ws.to_pdf(str(pdf_path))

                generated_files.append(str(pdf_path))
                logger.info(f"Successfully generated: {pdf_name}")

            except Exception as e:
                logger.error(f"Failed to generate PDF for sheet '{sheet_name}': {str(e)}")
                # Continue with other sheets instead of stopping
                continue

        wb.close()

        logger.info(f"PDF generation complete. Generated {len(generated_files)}/{len(sheet_names)} files")
        return generated_files

    except Exception as e:
        logger.error(f"PDF generation failed: {str(e)}")
        raise
    finally:
        app.quit()


if __name__ == "__main__":
    # Test code
    test_excel = "プログラム基本設計書_累積作成ツール.xlsx"
    if Path(test_excel).exists():
        sheets = list_all_sheets(test_excel)
        print(f"Available sheets: {sheets}")

        # Test PDF generation with first sheet
        if sheets:
            pdfs = generate_pdfs(test_excel, [sheets[0]], 6, "./output/pdfs")
            print(f"Generated PDFs: {pdfs}")
