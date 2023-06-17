import math
from num_to_bd import num_bd_dict, num_bd_letter_dict, division_words


class EnToBnValue:

    def __init__(self, value) -> None:
        self.number = value
        self.words = ''

    def in_bn_letter(self):
        number_str = str(self.number)
        if "." in number_str:
            decimal_list = number_str.split(".")
            before_dec = [num_bd_letter_dict[i] for i in decimal_list[0]]
            after_dec = [num_bd_letter_dict[i] for i in decimal_list[1]]
            return ''.join(before_dec)+'.'+''.join(after_dec)
        else:
            number_bn_list = [num_bd_letter_dict[i] for i in number_str]
            return ''.join(number_bn_list)

    def three_digit(self, three_digits):
        in_words = ''
        if three_digits[0] != "0":
            in_words += f'{num_bd_dict[three_digits[:-2]]}শত '
        in_words += num_bd_dict[three_digits[-2:]]
        return in_words

    def in_words(self):
        number_str = str(self.number)
        if len(number_str) > 3:
            last_three = number_str[-3:]
            rest = number_str[:-3]
            self.words += self.three_digit(last_three)
            i = 0
            x = -2
            while i < (math.ceil(len(rest)/2)):
                if i == 0:
                    if rest[x:] != "00":
                        self.words = f'{num_bd_dict[rest[x:]]} {division_words[i]} '+self.words
                else:
                    if rest[x:x+2] != "00":
                        self.words = f'{num_bd_dict[rest[x:x+2]]} {division_words[i]} '+self.words
                i += 1
                x -= 2
        else:
            self.words = self.three_digit(number_str)

        return self.words
