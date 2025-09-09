import os
import platform
import shutil
import subprocess
import time
import numpy as np
import cv2 as cv
import fitz
from pdf2docx import Converter


# === SET YOUR INPUT PDF FILE PATH HERE ===
input_pdf = r"18.pdf"  # Change this path as needed

# Output filenames (saved in current directory)
output_docx = "DT0122_U1800_TDS-P15B_R0.docx"

def run(command):
    print(f'Running: {command}')
    subprocess.run(command, shell=True, check=True)


def document_to(in_, out):
    """Convert DOCX to PDF using Word on Windows or LibreOffice on other platforms."""
    if platform.system() == 'Windows':
        word_to(in_, out)
    else:
        libreoffice_to(in_, out)


def word_to(in_, out):
    import docx2pdf
    assert os.path.isfile(in_), f'Input file missing: {in_}'
    try:
        docx2pdf.convert(in_, out)
    except Exception as e:
        print(f'docx2pdf conversion error: {e}')
        raise


def libreoffice_to(in_, out):
    assert os.path.isfile(in_), f'Input file missing: {in_}'
    out_dir = os.path.dirname(out)
    in_root, in_ext = os.path.splitext(in_)
    _, out_ext = os.path.splitext(out)
    temp_dir = os.path.join(out_dir, '_temp_libreoffice_to')
    os.makedirs(temp_dir, exist_ok=True)

    temp_in = os.path.join(temp_dir, f'temp{in_ext}')
    temp_out = os.path.join(temp_dir, f'temp{out_ext}')
    shutil.copy2(in_, temp_in)

    try:
        t = time.time()
        run(f'libreoffice --headless --convert-to {out_ext[1:]} --outdir {temp_dir} "{temp_in}"')
        os.rename(temp_out, out)
        t_out = os.path.getmtime(out)
        assert t_out >= t, f'LibreOffice conversion failed to create {out}'
    finally:
        os.remove(temp_in)
        if os.path.isfile(temp_out):
            os.remove(temp_out)


def get_page_image(pdf_page):
    """Convert fitz page to OpenCV image."""
    pix = pdf_page.get_pixmap()
    img_bytes = pix.tobytes()
    img_array = np.frombuffer(img_bytes, np.uint8)
    img = cv.imdecode(img_array, cv.IMREAD_COLOR)
    return img


def get_mssism(i1, i2, kernel=(15, 15)):
    """Calculate mean SSIM similarity score between two images."""
    C1 = 6.5025
    C2 = 58.5225
    I1 = np.float32(i1)
    I2 = np.float32(i2)
    I2_2 = I2 * I2
    I1_2 = I1 * I1
    I1_I2 = I1 * I2

    mu1 = cv.GaussianBlur(I1, kernel, 1.5)
    mu2 = cv.GaussianBlur(I2, kernel, 1.5)
    mu1_2 = mu1 * mu1
    mu2_2 = mu2 * mu2
    mu1_mu2 = mu1 * mu2

    sigma1_2 = cv.GaussianBlur(I1_2, kernel, 1.5) - mu1_2
    sigma2_2 = cv.GaussianBlur(I2_2, kernel, 1.5) - mu2_2
    sigma12 = cv.GaussianBlur(I1_I2, kernel, 1.5) - mu1_mu2

    t1 = 2 * mu1_mu2 + C1
    t2 = 2 * sigma12 + C2
    t3 = t1 * t2

    t1 = mu1_2 + mu2_2 + C1
    t2 = sigma1_2 + sigma2_2 + C2
    t1 = t1 * t2

    ssim_map = cv.divide(t3, t1)
    mssim = cv.mean(ssim_map)
    return np.mean(mssim[:3])


def get_page_similarity(page_a, page_b):
    """Calculate similarity index [0,1] between two PDF pages."""
    img_a = get_page_image(page_a)
    img_b = get_page_image(page_b)

    if img_a.shape != img_b.shape:
        img_b = cv.resize(img_b, (img_a.shape[1], img_a.shape[0]))

    return get_mssism(img_a, img_b)


def compare_pdf(pdf1, pdf2):
    """Compare two PDFs page by page, return average similarity."""
    with fitz.open(pdf1) as doc1, fitz.open(pdf2) as doc2:
        if len(doc1) != len(doc2):
            print(f'Page count mismatch: {len(doc1)} vs {len(doc2)}')
            return -1

        total_similarity = 0.0
        for i in range(len(doc1)):
            sim = get_page_similarity(doc1[i], doc2[i])
            print(f'Page {i+1} similarity: {sim:.4f}')
            total_similarity += sim

        return total_similarity / len(doc1)


def convert_pdf_to_docx(pdf_file, docx_file):
    """Convert PDF to DOCX using pdf2docx Converter."""
    c = Converter(pdf_file)
    c.convert(docx_file)
    c.close()


def convert_docx_to_pdf(docx_file, pdf_file):
    """Convert DOCX back to PDF using Word or LibreOffice."""
    document_to(docx_file, pdf_file)


def test_conversion():
    """Full pipeline: PDF -> DOCX -> PDF and compare."""
    print(f'Converting {input_pdf} to DOCX...')
    convert_pdf_to_docx(input_pdf, output_docx)
    assert os.path.isfile(output_docx), "DOCX conversion failed."

    print(f'Converting {output_docx} back to PDF...')
    convert_docx_to_pdf(output_docx, output_pdf)
    assert os.path.isfile(output_pdf), "PDF reconversion failed."

    print('Comparing original and reconverted PDFs...')
    similarity = compare_pdf(input_pdf, output_pdf)
    print(f'Average page similarity: {similarity:.4f}')

    # Optional threshold
    threshold = 0.85
    assert similarity >= threshold, f'Similarity below threshold {threshold}'


if __name__ == "__main__":
    try:
        test_conversion()
        print("\n✅ Conversion and comparison successful.")
        print(f"Saved as:\n- {output_docx}")
    except Exception as e:
        print(f"\n❌ Error: {e}")


