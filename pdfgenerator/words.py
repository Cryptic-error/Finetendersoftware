def rupee_format_get_d(x_str):
    arr_1 = ["One", "Two", "Three", "Four", "Five", "Six", "Seven", "Eight", "Nine", ""]
    return arr_1[int(x_str) - 1] if int(x_str) > 0 else ""

def rupee_format_get_t(x_str):
    arr1 = ["Ten", "Eleven", "Twelve", "Thirteen", "Fourteen", "Fifteen", "Sixteen", "Seventeen", "Eighteen", "Nineteen"]
    arr2 = ["", "Twenty", "Thirty", "Forty", "Fifty", "Sixty", "Seventy", "Eighty", "Ninety"]
    result = ""

    if int(x_str[0]) == 1:
        result = arr1[int(x_str[1])]
    else:
        if int(x_str[0]) > 0:
            result = arr2[int(x_str[0]) - 1]
        result += rupee_format_get_d(x_str[1])

    return result

def rupee_format_get_h(x_str_h, x_lp):
    x_str_h = x_str_h.zfill(3)  # Pad to ensure it's 3 digits
    result = ""

    if int(x_str_h) < 1:
        return ""

    if x_str_h[0] != "0":
        if x_lp > 0:
            result = rupee_format_get_d(x_str_h[0]) + " Lac "
        else:
            result = rupee_format_get_d(x_str_h[0]) + " Hundred "

    if x_str_h[1] != "0":
        result += rupee_format_get_t(x_str_h[1:])
    else:
        result += rupee_format_get_d(x_str_h[2])

    return result

def rupee_format(s_num):
    arr_place = ["", "", " Thousand ", " Lacs ", " Crores ", " Trillion ", "", "", "", ""]
    if s_num == "":
        return ""

    x_num_str = str(s_num).strip()

    if x_num_str == "":
        return ""

    if float(x_num_str) > 999999999.99:
        return "Digit exceeds Maximum limit"

    x_dp_int = x_num_str.find(".")

    x_rstr_paisas = ""  # Initialize x_rstr_paisas here

    if x_dp_int > 0:
        if len(x_num_str) - x_dp_int == 1:
            x_rstr_paisas = rupee_format_get_t(x_num_str[x_dp_int + 1:x_dp_int + 2] + "0")
        elif len(x_num_str) - x_dp_int > 1:
            x_rstr_paisas = rupee_format_get_t(x_num_str[x_dp_int + 1:x_dp_int + 3])

        x_num_str = x_num_str[:x_dp_int]

    x_f = 1
    x_rstr = ""
    x_lp = 0

    while x_num_str != "":
        if x_f >= 2:
            x_temp = x_num_str[-2:]
        else:
            if len(x_num_str) == 2:
                x_temp = x_num_str[-2:]
            elif len(x_num_str) == 1:
                x_temp = x_num_str[-1:]
            else:
                x_temp = x_num_str[-3:]

        x_str_temp = ""

        if int(x_temp) > 99:
            x_str_temp = rupee_format_get_h(x_temp, x_lp)
            if "Lac" not in x_str_temp:
                x_lp += 1
        elif int(x_temp) <= 99 and int(x_temp) > 9:
            x_str_temp = rupee_format_get_t(x_temp)
        elif int(x_temp) < 10:
            x_str_temp = rupee_format_get_d(x_temp)

        if x_str_temp != "":
            x_rstr = x_str_temp + arr_place[x_f] + x_rstr

        if x_f == 2:
            if len(x_num_str) == 1:
                x_num_str = ""
            else:
                x_num_str = x_num_str[:-2]
        elif x_f == 3:
            if len(x_num_str) >= 3:
                x_num_str = x_num_str[:-2]
            else:
                x_num_str = ""
        elif x_f == 4:
            x_num_str = ""
        else:
            if len(x_num_str) <= 2:
                x_num_str = ""
            else:
                x_num_str = x_num_str[:-3]

        x_f += 1

    if x_rstr == "":
        x_rstr = "No Rupees"
    else:
        x_rstr = "Rupees " + x_rstr

    if x_rstr_paisas != "":
        x_rstr_paisas = " and " + x_rstr_paisas + " Paisas"

    return x_rstr + x_rstr_paisas + " Only"

# Example Usage:
amt = (5000)

if __name__ == "__main__":
    print(rupee_format(amt)) 
    print(type(amt))         # "Rupees Five Thousand Only"
    print(rupee_format("1234567.89"))   # "Rupees Twelve Lacs Thirty Four Thousand Five Hundred Sixty Seven and Eighty Nine Paisas Only"
    print(rupee_format("1000000000"))   # "Digit exceeds Maximum limit"
