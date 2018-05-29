# -*- coding: utf-8 -*-


def render(digits):
    """
    This function converts its input: a string of decimal digits into a barcode
    using the interleaved 2 of 5 or 128 format
    The input string must not contain an odd number of digits. The output is an SVG string.
    :param digits: Number to translate
    :return: SVG string Representing a barcode
    """
    top = '<svg height="250" width="{0}" style="background:white">'
    bar = '<rect x="{0}" y="120.6" width="{1}" height="32.2" style="fill:black"/>'
    barcode = [top.format(len(digits) * 14 + 24)]

    def checksum():
        """
        Perform the UPC checksum of the digits encoded and add it to the code
        :return: -
        """
        acc = 0
        for i in range(len(digits)):
            if i % 2 == 0:
                acc += 3 * int(digits[i])
            else:
                acc += int(digits[i])
        m = acc % 10
        if m != 0:
            return str(10-m)
        return '0'

    def itof(num):
        """
        Switch implementation matching a number to its byte representation (length of rectangles for itof barcode)
        :param num: The index of digits to convert
        :return: Its rectangles length for barcode representation
        """
        return {'0': (1.2, 1.2, 2.4, 2.4, 1.2), '1': (2.4, 1.2, 1.2, 1.2, 2.4), '2': (1.2, 2.4, 1.2, 1.2, 2.4),
                '3': (2.4, 2.4, 1.2, 1.2, 1.2), '4': (1.2, 1.2, 2.4, 1.2, 2.4), '5': (2.4, 1.2, 2.4, 1.2, 1.2),
                '6': (1.2, 2.4, 2.4, 1.2, 1.2), '7': (1.2, 1.2, 1.2, 2.4, 2.4), '8': (2.4, 1.2, 1.2, 2.4, 1.2),
                '9': (1.2, 2.4, 1.2, 2.4, 1.2)}[digits[num]]

    def encode(bits, i=0):
        """
        Encodes a string of decimal digits into a list of bar lengths for the barcode (black and white ones)
        :param bits: List of ints already encode (white and black bar representations)
        :param i: The index of the element of digits to encode
        :return: Representation of digits as a list of ints (with length for barcode rectangles)
        """
        if i == len(digits):
            return bits

        for b, w in zip(itof(i), itof(i+1)):
            bits.extend([b, w])

        return encode(bits, i+2)

    def svg(bits, i=0, x=22):
        """
        Converts the list of ints returned by `encode` to a list of SVG
        elements that can be joined to form an SVG string
        :param bits: Representation of digits as a list of ints (with length for barcode rectangles)
        :param i: Index of bit to transform into svg lines
        :param x: Place on which the rectangle (representing a bar) should be placed in the svg
        :return: SVG string representation of the digits
        """
        if i == len(bits):
            return barcode + ['</svg>']

        b, w = bits[i:i+2]
        barcode.append(bar.format(x, b))

        return svg(bits, i+2, x+b+w)

    digits += checksum()
    return '\n'.join(svg(encode([1.2, 1.2, 1.2, 1.2]) + [2.4, 1.2, 1.2, 0]))
