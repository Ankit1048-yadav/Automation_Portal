import os
import cv2
import numpy as np

from PIL import Image, ImageEnhance, ImageFilter

# ===========================
# A4 SIZE @300 DPI
# ===========================

PAGE_WIDTH = 2480
PAGE_HEIGHT = 3508

MARGIN = 40

# ===========================
# IMAGE READ
# ===========================

def read_image(path):

    img = cv2.imread(path)

    if img is None:
        raise Exception(f"Cannot read image : {path}")

    return img


# ===========================
# ORDER 4 CORNERS
# ===========================

def order_points(pts):

    rect = np.zeros((4,2),dtype="float32")

    s = pts.sum(axis=1)

    rect[0] = pts[np.argmin(s)]
    rect[2] = pts[np.argmax(s)]

    diff = np.diff(pts,axis=1)

    rect[1] = pts[np.argmin(diff)]
    rect[3] = pts[np.argmax(diff)]

    return rect


# ===========================
# PERSPECTIVE TRANSFORM
# ===========================

def four_point_transform(image,pts):

    rect = order_points(pts)

    (tl,tr,br,bl)=rect

    widthA=np.linalg.norm(br-bl)
    widthB=np.linalg.norm(tr-tl)

    maxWidth=max(int(widthA),int(widthB))

    heightA=np.linalg.norm(tr-br)
    heightB=np.linalg.norm(tl-bl)

    maxHeight=max(int(heightA),int(heightB))

    dst=np.array([
        [0,0],
        [maxWidth-1,0],
        [maxWidth-1,maxHeight-1],
        [0,maxHeight-1]
    ],dtype="float32")

    M=cv2.getPerspectiveTransform(rect,dst)

    warped=cv2.warpPerspective(
        image,
        M,
        (maxWidth,maxHeight)
    )

    return warped


# ===========================
# FIND DOCUMENT
# ===========================

def detect_document(img):

    ratio=img.shape[0]/500.0

    original=img.copy()

    resized=cv2.resize(
        img,
        (
            int(img.shape[1]/ratio),
            500
        )
    )

    gray=cv2.cvtColor(
        resized,
        cv2.COLOR_BGR2GRAY
    )

    gray=cv2.GaussianBlur(
        gray,
        (5,5),
        0
    )

    edged=cv2.Canny(
        gray,
        30,
        120
    )

    contours,_=cv2.findContours(
        edged,
        cv2.RETR_LIST,
        cv2.CHAIN_APPROX_SIMPLE
    )

    contours=sorted(
        contours,
        key=cv2.contourArea,
        reverse=True
    )[:10]

    screenCnt=None

    for c in contours:

        peri=cv2.arcLength(c,True)

        approx=cv2.approxPolyDP(
            c,
            0.02*peri,
            True
        )

        if len(approx)==4:

            screenCnt=approx

            break

    if screenCnt is None:

        return original

    warped=four_point_transform(
        original,
        screenCnt.reshape(4,2)*ratio
    )

    return warped

# ===========================
# IMAGE ENHANCEMENT
# ===========================

def enhance_document(img):

    # OpenCV -> PIL
    rgb = cv2.cvtColor(img, cv2.COLOR_BGR2RGB)

    pil = Image.fromarray(rgb)

    # Contrast
    pil = ImageEnhance.Contrast(pil).enhance(1.10)

    # Brightness
    pil = ImageEnhance.Brightness(pil).enhance(1.02)

    # Sharpness
    pil = ImageEnhance.Sharpness(pil).enhance(1.3)

    # Smooth
    pil = pil.filter(ImageFilter.MedianFilter(size=3))

    return pil


# ===========================
# REMOVE BLACK BORDER
# ===========================

def auto_crop_white(pil):

    img = np.array(pil)

    gray = cv2.cvtColor(img, cv2.COLOR_RGB2GRAY)

    thresh = cv2.adaptiveThreshold(
    gray,
    255,
    cv2.ADAPTIVE_THRESH_GAUSSIAN_C,
    cv2.THRESH_BINARY_INV,
    31,
    15
    )

    kernel = np.ones((7,7), np.uint8)

    thresh = cv2.morphologyEx(
        thresh,
        cv2.MORPH_CLOSE,
        kernel,
        iterations=2
    )

    contours, _ = cv2.findContours(
        thresh,
        cv2.RETR_EXTERNAL,
        cv2.CHAIN_APPROX_SIMPLE
    )

    if len(contours) == 0:
        return pil

    c = max(contours, key=cv2.contourArea)

    x, y, w, h = cv2.boundingRect(c)

    pad = 5

    x = max(0, x - pad)
    y = max(0, y - pad)

    w = min(img.shape[1] - x, w + pad * 2)
    h = min(img.shape[0] - y, h + pad * 2)

    crop = pil.crop((x, y, x + w, y + h))

    return crop
# =====================================================
# AUTO FIT DOCUMENT
# =====================================================

def auto_fit_document(pil):

    max_width = PAGE_WIDTH - (2 * MARGIN)
    max_height = PAGE_HEIGHT - (2 * MARGIN)

    ratio = min(
        max_width / pil.width,
        max_height / pil.height
    )

    # Document ko page ka ~98% tak use karne do
    ratio *= 0.98

    new_width = int(pil.width * ratio)
    new_height = int(pil.height * ratio)

    return pil.resize(
        (new_width, new_height),
        Image.LANCZOS
    )

# ===========================
# FIT TO A4 PAGE
# ===========================

def create_a4_page(pil):

    canvas = Image.new(
        "RGB",
        (PAGE_WIDTH, PAGE_HEIGHT),
        "white"
    )

    resized = auto_fit_document(pil)

    new_w = resized.width
    new_h = resized.height

    x = (PAGE_WIDTH - new_w) // 2
    y = (PAGE_HEIGHT - new_h) // 2

    canvas.paste(
        resized,
        (x, y)
    )

    return canvas

# ===========================
# GENERATE PDF
# ===========================

def generate_document_pdf(image_paths, output_pdf):

    pages = []

    if len(image_paths) == 0:
        raise Exception("No images selected.")

    for path in image_paths:

        print("Processing :", path)

        try:

            page = process_single_image(path)

            pages.append(page.convert("RGB"))

        except Exception as e:

            print(f"Skipped : {path}")
            print(e)

    if len(pages) == 0:

        raise Exception("No valid images found.")

    pages[0].save(
        output_pdf,
        "PDF",
        resolution=300.0,
        save_all=True,
        append_images=pages[1:]
    )

    print("=================================")
    print("PDF Generated Successfully")
    print(output_pdf)
    print("=================================")

    return output_pdf

# ===========================
# OVERRIDE PROCESS FUNCTION
# ===========================

def process_single_image(path):

    img = read_image(path)

    img = auto_rotate(img)

    doc = detect_document(img)

    enhanced = enhance_document(doc)

    enhanced = auto_crop_white(enhanced)

    page = create_a4_page(enhanced)

    return page
# =====================================================
# AUTO WHITE BALANCE
# =====================================================

def auto_white_balance(img):

    lab = cv2.cvtColor(
        img,
        cv2.COLOR_BGR2LAB
    )

    l, a, b = cv2.split(lab)

    l = cv2.equalizeHist(l)

    lab = cv2.merge((l, a, b))

    result = cv2.cvtColor(
        lab,
        cv2.COLOR_LAB2BGR
    )

    return result


# =====================================================
# SHADOW REMOVAL
# =====================================================

def remove_shadow(img):

    rgb_planes = cv2.split(img)

    result_planes = []

    for plane in rgb_planes:

        dilated = cv2.dilate(
            plane,
            np.ones((7,7), np.uint8)
        )

        bg = cv2.medianBlur(
            dilated,
            21
        )

        diff = 255 - cv2.absdiff(
            plane,
            bg
        )

        norm = cv2.normalize(
            diff,
            None,
            0,
            255,
            cv2.NORM_MINMAX,
            dtype=cv2.CV_8UC1
        )

        result_planes.append(norm)

    result = cv2.merge(result_planes)

    return result


# =====================================================
# COLOR ENHANCEMENT
# =====================================================

def enhance_color(img):

    img = auto_white_balance(img)

    img = remove_shadow(img)

    img = cv2.detailEnhance(
        img,
        sigma_s=10,
        sigma_r=0.15
    )

    img = cv2.bilateralFilter(
        img,
        9,
        75,
        75
    )

    return img


# =====================================================
# AUTO ROTATE
# =====================================================

def auto_rotate(img):

    h, w = img.shape[:2]

    if h > w * 1.8:

        return img

    if w > h * 1.3:

        return cv2.rotate(
            img,
            cv2.ROTATE_90_CLOCKWISE
        )

    return img

# =====================================================
# SMART DOCUMENT ENHANCEMENT
# =====================================================

def smart_document_enhancement(img):

    gray = cv2.cvtColor(
        img,
        cv2.COLOR_BGR2GRAY
    )

    gray = cv2.fastNlMeansDenoising(
        gray,
        None,
        8,
        7,
        21
    )

    clahe = cv2.createCLAHE(
        clipLimit=2.0,
        tileGridSize=(8,8)
    )

    gray = clahe.apply(gray)

    rgb = cv2.cvtColor(
        gray,
        cv2.COLOR_GRAY2BGR
    )

    return rgb
# =====================================================
# FINAL IMAGE TUNING
# =====================================================

def final_image_tuning(img):

    gray = cv2.cvtColor(
        img,
        cv2.COLOR_BGR2GRAY
    )

    # Mild denoising
    gray = cv2.GaussianBlur(
        gray,
        (3,3),
        0
    )

    # Improve local contrast
    clahe = cv2.createCLAHE(
        clipLimit=1.8,
        tileGridSize=(8,8)
    )

    gray = clahe.apply(gray)

    # Light sharpening
    kernel = np.array([
        [0,-1,0],
        [-1,5,-1],
        [0,-1,0]
    ])

    gray = cv2.filter2D(
        gray,
        -1,
        kernel
    )

    return cv2.cvtColor(
        gray,
        cv2.COLOR_GRAY2BGR
    )