# -*- coding: utf-8 -*-


def render(digits):
    """""
        This function converts its input, a string of decimal digits, into a
        barcode, using the interleaved 2 of 5 format. The input string must not
        contain an odd number of digits. The output is an SVG string.

        Wikipedia [ITF Format]: http://en.wikipedia.org/wiki/Interleaved_2_of_5
    """

    top = '<svg height="250" width="{0}" style="background:white">'
    bar = '<rect x="{0}" y="120.6" width="{1}" height="32.2" style="fill:black"/>'
    barcode = [top.format(len(digits) * 14 + 24)]

    def byte(num):
        return {'0': (1.4, 1.4, 2.8, 2.8, 1.4), '1': (2.8, 1.4, 1.4, 1.4, 2.8), '2': (1.4, 2.8, 1.4, 1.4, 2.8),
                '3': (2.8, 2.8, 1.4, 1.4, 1.4), '4': (1.4, 1.4, 2.8, 1.4, 2.8), '5': (2.8, 1.4, 2.8, 1.4, 1.4),
                '6': (1.4, 2.8, 2.8, 1.4, 1.4), '7': (1.4, 1.4, 1.4, 2.8, 2.8), '8': (2.8, 1.4, 1.4, 2.8, 1.4),
                '9': (1.4, 2.8, 1.4, 2.8, 1.4)}[digits[num]]

    def encode(bits, i=0):

        """Encodes a string of decimal digits. The output's a binary string,
        but as a list of ints, all 2s and 4s, rather than zeros and ones, as
        the bars in the SVG barcode are rendered 2px wide for zeros, 4px for
        ones."""

        if i == len(digits):
            return bits

        for b, w in zip(byte(i), byte(i+1)):
            bits.extend([b, w])

        return encode(bits, i+2)

    def svg(bits, i=0, x=10):

        """Converts the list of ints returned by `encode` to a list of SVG
        elements that can be joined to form an SVG string."""

        if i == len(bits):
            return barcode + ['</svg>']

        b, w = bits[i:i+2]
        barcode.append(bar.format(x, b))

        return svg(bits, i+2, x+b+w)

    # The call to `encode` passes in the itf start code, and the itf end
    # code's added to the list that `encode` returns. The complete list
    # is passed to `svg`, which returns a list of SVG element strings
    # that're joined with newlines to create the actual SVG output.

    return '\n'.join(svg(encode([2, 2, 2, 2]) + [4, 2, 2, 0]))