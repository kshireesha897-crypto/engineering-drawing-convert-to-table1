import pytesseract

def GetString(img, word1, word2):

    text = pytesseract.image_to_string(img)

    lines = text.splitlines()

    for line in lines:

        if word1 in line or word2 in line:
            return line

    return ""