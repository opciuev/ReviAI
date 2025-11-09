"""
Step 2: AI Review module using Gemini 2.5 Pro
"""
import warnings
# Suppress warnings from dependencies
warnings.filterwarnings("ignore", message="Unsupported Windows version")
warnings.filterwarnings("ignore", message="Couldn't find ffmpeg or avconv")
warnings.filterwarnings("ignore", category=UserWarning, module="onnxruntime")
warnings.filterwarnings("ignore", category=RuntimeWarning, module="pydub")

from google import genai
from google.genai import types
from pathlib import Path
from typing import List
from models import ReviewTable
from logger import logger
from markitdown import MarkItDown
import json
from datetime import datetime


def convert_files_to_markdown(file_paths: List[str]) -> str:
    """
    Convert multiple PDF or Markdown files to combined Markdown format

    Args:
        file_paths: List of PDF or Markdown file paths

    Returns:
        str: Combined Markdown content from all files

    Raises:
        FileNotFoundError: If any file doesn't exist
        Exception: If conversion fails
    """
    md_converter = MarkItDown()
    combined_markdown = []

    for idx, file_path in enumerate(file_paths, 1):
        if not Path(file_path).exists():
            raise FileNotFoundError(f"File not found: {file_path}")

        filename = Path(file_path).name
        file_ext = Path(file_path).suffix.lower()

        logger.info(f"Processing file: {filename}")

        try:
            # Check if it's a Markdown file
            if file_ext == '.md':
                # Read Markdown file directly
                with open(file_path, 'r', encoding='utf-8') as f:
                    markdown_content = f.read()
                logger.info(f"Read Markdown file: {filename}")
            else:
                # Convert PDF to Markdown
                result = md_converter.convert(file_path)
                markdown_content = result.text_content
                logger.info(f"Converted PDF to Markdown: {filename}")

            # Detect version from filename (V6, V7, etc.)
            version_marker = ""
            if "V6" in filename.upper() or "_6" in filename:
                version_marker = " (前回の設計書 V6)"
            elif "V7" in filename.upper() or "_7" in filename:
                version_marker = " (今回の設計書 V7)"
            else:
                version_marker = f" (Document {idx})"

            # Create clear document boundary with version info
            file_section = f"""
{'='*80}
ドキュメント: {Path(file_path).stem}{file_ext}{version_marker}
ファイル名: {filename}
{'='*80}

{markdown_content}
"""
            combined_markdown.append(file_section)
            logger.info(f"Successfully processed: {filename}{version_marker}")

        except Exception as e:
            logger.error(f"Failed to process {file_path}: {str(e)}")
            raise Exception(f"File processing failed for {filename}: {str(e)}")

    # Combine all markdown content with clear separators
    full_markdown = "\n\n".join(combined_markdown)
    logger.info(f"Combined {len(file_paths)} files into Markdown ({len(full_markdown)} characters)")

    return full_markdown


def convert_pdfs_to_markdown(pdf_paths: List[str]) -> str:
    """
    Legacy function for backward compatibility.
    Redirects to convert_files_to_markdown.
    """
    return convert_files_to_markdown(pdf_paths)


def review_with_gemini(
    pdf_paths: List[str],
    prompt_template: str,
    api_key: str,
    model: str = "gemini-2.5-pro"
) -> ReviewTable:
    """
    Perform AI review using Gemini Pro

    Args:
        pdf_paths: List of PDF or Markdown file paths to review
        prompt_template: Prompt template content
        api_key: Gemini API key
        model: Model name (default: gemini-2.5-pro)

    Returns:
        ReviewTable: Structured review results

    Raises:
        FileNotFoundError: If file doesn't exist
        ValueError: If API key is invalid
        Exception: If API call fails
    """
    # Validate API key
    if not api_key or api_key == "YOUR_API_KEY_HERE" or api_key == "YOUR_GEMINI_API_KEY_HERE":
        raise ValueError("Invalid API key. Please configure your Gemini API key in config.ini")

    # Validate files exist
    for file_path in pdf_paths:
        if not Path(file_path).exists():
            raise FileNotFoundError(f"File not found: {file_path}")

    logger.info(f"Starting AI review with {len(pdf_paths)} files")
    logger.info(f"Using model: {model}")

    try:
        # Convert files to Markdown
        logger.info("Processing files to Markdown format...")
        markdown_content = convert_files_to_markdown(pdf_paths)
        logger.info(f"File processing completed. Total content length: {len(markdown_content)} characters")

        # Create debug directory
        debug_dir = Path("./output/debug")
        debug_dir.mkdir(parents=True, exist_ok=True)
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")

        # Save Markdown content
        markdown_file = debug_dir / f"pdf_markdown_{timestamp}.md"
        with open(markdown_file, 'w', encoding='utf-8') as f:
            f.write(markdown_content)
        logger.info(f"Saved Markdown content to: {markdown_file}")

        # Initialize Gemini client
        client = genai.Client(api_key=api_key)

        # Build request contents (prompt + markdown content)
        full_prompt = f"""{prompt_template}

# PDF Content (Markdown format):

{markdown_content}
"""

        # Save full prompt
        prompt_file = debug_dir / f"full_prompt_{timestamp}.txt"
        with open(prompt_file, 'w', encoding='utf-8') as f:
            f.write(full_prompt)
        logger.info(f"Saved full prompt to: {prompt_file}")

        logger.info("Calling Gemini API for review...")

        # Call Gemini API with JSON Schema
        response = client.models.generate_content(
            model=model,
            contents=full_prompt,
            config=types.GenerateContentConfig(
                response_mime_type='application/json',
                response_schema=ReviewTable,
                temperature=0,
                max_output_tokens=65536,  # Increased from 8192 to support large tables
            )
        )

        # Save raw API response (if available)
        try:
            raw_response_file = debug_dir / f"api_response_raw_{timestamp}.txt"
            with open(raw_response_file, 'w', encoding='utf-8') as f:
                f.write(f"Model: {model}\n")
                f.write(f"Response text: {response.text if hasattr(response, 'text') else 'N/A'}\n")
                f.write(f"Candidates: {len(response.candidates) if hasattr(response, 'candidates') else 0}\n")
                if hasattr(response, 'usage_metadata'):
                    f.write(f"Usage metadata: {response.usage_metadata}\n")
            logger.info(f"Saved raw API response to: {raw_response_file}")
        except Exception as e:
            logger.warning(f"Could not save raw response: {e}")

        # Get parsed structured data
        result: ReviewTable = response.parsed

        if result is None:
            logger.error("Failed to parse API response")
            raise Exception("Gemini API returned None for parsed result")

        # Save parsed result as JSON
        result_file = debug_dir / f"parsed_result_{timestamp}.json"
        with open(result_file, 'w', encoding='utf-8') as f:
            json.dump(result.model_dump(), f, ensure_ascii=False, indent=2)
        logger.info(f"Saved parsed result to: {result_file}")

        logger.info(f"AI review completed successfully. Extracted {len(result.rows)} rows")
        logger.info(f"⚠️ DEBUG: All debug files saved to: {debug_dir}")

        return result

    except Exception as e:
        logger.error(f"AI review failed: {str(e)}")
        raise


def review_with_retry(
    pdf_paths: List[str],
    prompt_template: str,
    api_key: str,
    model: str = "gemini-2.5-pro",
    max_retries: int = 3
) -> ReviewTable:
    """
    Perform AI review with retry logic

    Args:
        pdf_paths: List of PDF file paths
        prompt_template: Prompt template content
        api_key: Gemini API key
        model: Model name
        max_retries: Maximum number of retry attempts

    Returns:
        ReviewTable: Structured review results

    Raises:
        Exception: If all retry attempts fail
    """
    import time

    last_exception = None

    for attempt in range(max_retries):
        try:
            logger.info(f"Attempt {attempt + 1}/{max_retries}")
            result = review_with_gemini(pdf_paths, prompt_template, api_key, model)
            return result

        except Exception as e:
            last_exception = e
            logger.warning(f"Attempt {attempt + 1} failed: {str(e)}")

            if attempt < max_retries - 1:
                wait_time = 3 * (2 ** attempt)  # Exponential backoff: 3s, 6s, 12s
                logger.info(f"Retrying in {wait_time} seconds...")
                time.sleep(wait_time)
            else:
                logger.error(f"All {max_retries} attempts failed")

    raise last_exception


if __name__ == "__main__":
    # Test code
    from config_manager import ConfigManager

    # Load configuration
    config = ConfigManager.load_config()
    api_key = config['API']['gemini_api_key']
    prompt = ConfigManager.get_prompt_template()

    # Test with sample PDF (if exists)
    test_pdf = "output/pdfs/sample_V6.pdf"
    if Path(test_pdf).exists():
        try:
            result = review_with_gemini(
                [test_pdf],
                prompt,
                api_key
            )
            print(f"Review completed: {len(result.rows)} rows extracted")
            for i, row in enumerate(result.rows[:3], 1):
                print(f"\nRow {i}:")
                print(f"  要求No: {row.requirement_no}")
                print(f"  評価: {row.evaluation}")
        except Exception as e:
            print(f"Test failed: {str(e)}")
    else:
        print(f"Test PDF not found: {test_pdf}")
