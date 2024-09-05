from fractions import Fraction
import math
import re
import import_data


# length_input = str(input("Input length of pipe in format - {XX'XX-XX/XX\"}\n"))


coefficient_cs = 0.0000067


def read_length(length):
    true_length = length.replace('"', '')

    input_check = re.findall(r"\d+\'\d+\-\d+\/\d+", true_length)
    if input_check:

        # print("Input correct!")

        true_length = true_length.split("'")
        true_feet = float(true_length[0])
        true_inches = float(true_length[1].split("-")[0]) + float(int(true_length[1].split("-")[1].split("/")[0]) /
                                                                  int(true_length[1].split("-")[1].split("/")[1]))

        true_length = round(true_feet + (true_inches / 12), 4)

        # print("True length is - ", true_length, "\n")

        return true_length

    else:
        # print("Error - input is incorrect")
        return "Error - input is incorrect"


def calculate_max_expansion(length, operating_temp, installation_temp):
    max_operating_expansion = (operating_temp - installation_temp) * (length * 12) * coefficient_cs

    numerator = Fraction(max_operating_expansion).limit_denominator(16).as_integer_ratio()[0]
    denumerator = Fraction(max_operating_expansion).limit_denominator(16).as_integer_ratio()[1]

    coef_denum = round((16 / denumerator), 4)

    coef_numerator = math.ceil(round(numerator * coef_denum, 4))
    whole_part_num = coef_numerator // 16
    denumerator = 16

    answer = f"Max expansion - {round(whole_part_num, 0)}' {coef_numerator - (whole_part_num * 16)}/16\""

    clean_numerator = coef_numerator - (whole_part_num * 16)
    if clean_numerator == 0:
        answer = f"Max expansion - {round(whole_part_num, 0)}\' 0\""
    else:
        while clean_numerator % 2 == 0 and denumerator % 2 == 0:
            clean_numerator = int(round(clean_numerator / 2, 0))
            denumerator = int(round(denumerator / 2, 0))

            answer = f"Max expansion - {round(whole_part_num, 0)}' {clean_numerator}/{denumerator}\""

    return [whole_part_num, clean_numerator, denumerator, answer]


def calculate_max_compression(length, installation_temp, min_temp):
    max_operating_compression = (installation_temp - min_temp) * (length * 12) * coefficient_cs

    numerator = Fraction(max_operating_compression).limit_denominator(16).as_integer_ratio()[0]
    denumerator = Fraction(max_operating_compression).limit_denominator(16).as_integer_ratio()[1]

    coef_denum = round((16 / denumerator), 4)

    coef_numerator = math.ceil(round(numerator * coef_denum, 4))
    whole_part_num = coef_numerator // 16
    denumerator = 16

    answer = f"Max compression - {round(whole_part_num, 0)}' {coef_numerator - (whole_part_num * 16)}/16\""

    clean_numerator = coef_numerator - (whole_part_num * 16)
    if clean_numerator == 0:
        answer = f"Max compression - {round(whole_part_num, 0)}' 0\""
    else:
        while clean_numerator % 2 == 0 and denumerator % 2 == 0:
            clean_numerator = int(clean_numerator / 2)
            denumerator = int(denumerator / 2)

            answer = f"Max expansion - {round(whole_part_num, 0)}' {clean_numerator}/{denumerator}\""

    return [whole_part_num, clean_numerator, denumerator, answer]


def calculate_total_movements(expansion_measurements: list, compression_measurements: list):
    total_movements_inches = Fraction(expansion_measurements[1], expansion_measurements[2]) + \
                             Fraction(compression_measurements[1], compression_measurements[2])

    total_movements_feet = Fraction(expansion_measurements[0]) + Fraction(compression_measurements[0])

    whole_part_inches = 0
    if Fraction(total_movements_inches.numerator) >= Fraction(total_movements_inches.denominator):
        whole_part_inches = int(Fraction(total_movements_inches.numerator)) // \
                            int(Fraction(total_movements_inches.denominator))

    total_movements = f"{total_movements_feet}'{whole_part_inches}-{total_movements_inches}\""

    print("Total possible movements - ", total_movements, "\n")

    return total_movements


def pick_exj(expansion_measurements: list, compression_measurements: list,
             total_movements, exj_chart: dict, diameter_pipe):
    ls_chart = import_data.exj_long_stroke
    ss_chart = import_data.exj_short_stroke

    max_contraction_exj = Fraction(exj_chart[diameter_pipe][0].split(" ")[0]) + \
                          Fraction(exj_chart[diameter_pipe][0].split(" ")[1])
    max_expansion_exj = Fraction(exj_chart[diameter_pipe][1].split(" ")[0]) + \
                        Fraction(exj_chart[diameter_pipe][1].split(" ")[1])
    free_length_exj = Fraction(exj_chart[diameter_pipe][2].split(" ")[0]) + \
                      Fraction(exj_chart[diameter_pipe][2].split(" ")[1])

    expansion_value = "0/2"
    contraction_value = "0/2"

    if expansion_measurements[0] == 0:
        expansion_value = Fraction(expansion_measurements[1], expansion_measurements[2])
    else:
        expansion_value = Fraction(
            ((expansion_measurements[0] * expansion_measurements[2]) + expansion_measurements[1]),
            expansion_measurements[2])

    if compression_measurements[0] == 0:
        contraction_value = Fraction(compression_measurements[1], compression_measurements[2])
    else:
        contraction_value = Fraction(
            ((compression_measurements[0] * compression_measurements[2]) + compression_measurements[1]),
            compression_measurements[2])

    print("exp value - ", expansion_value)
    print("con value - ", contraction_value)

    min_length_exj = Fraction(free_length_exj) - Fraction(max_contraction_exj) + Fraction("1/4")
    max_length_exj = Fraction(free_length_exj) + Fraction(max_expansion_exj) - Fraction("1/4")

    print(f"min length exj - ", min_length_exj)
    print(f"max length exj - ", max_length_exj)

    check_exp_pipe_vs_exj = 0
    check_con_pipe_vs_exj = 0

    if expansion_measurements[0] == 0:
        if (expansion_measurements[1] / expansion_measurements[2]) <= max_contraction_exj:
            check_exp_pipe_vs_exj = 1
    else:
        if (expansion_measurements[0] * expansion_measurements[1] / expansion_measurements[2]) <= max_contraction_exj:
            check_exp_pipe_vs_exj = 1

    if compression_measurements[0] == 0:
        if (compression_measurements[1] / compression_measurements[2]) <= max_contraction_exj:
            check_con_pipe_vs_exj = 1
    else:
        if (compression_measurements[0] * compression_measurements[1] / compression_measurements[
            2]) <= max_contraction_exj:
            check_con_pipe_vs_exj = 1

    if check_con_pipe_vs_exj == 1 and check_exp_pipe_vs_exj == 1:
        print("\nMax. cont. EXJ > max. exp. PIPE & max. exp. EXJ > max. cont. PIPE - OK\n")
    else:
        print(f"\nSomething wrong - \n, CH_EXP -  {check_exp_pipe_vs_exj}, CH_CON - {check_con_pipe_vs_exj}")

    # suggest setting for EXJ

    suggested_setting_exj = ((Fraction(max_length_exj) + Fraction(contraction_value)) -
                             (Fraction(min_length_exj) + Fraction(expansion_value))) / 2
    suggested_setting_exj = suggested_setting_exj + Fraction(min_length_exj)

    print(f"Suggested EXJ setting - {suggested_setting_exj}")

    return suggested_setting_exj


def check_user_setting(expansion_measurements: list, compression_measurements: list,
                       total_movements, exj_chart: dict, diameter_pipe, user_setting):

    ls_chart = import_data.exj_long_stroke
    ss_chart = import_data.exj_short_stroke

    max_contraction_exj = Fraction(exj_chart[diameter_pipe][0].split(" ")[0]) + \
                          Fraction(exj_chart[diameter_pipe][0].split(" ")[1])
    max_expansion_exj = Fraction(exj_chart[diameter_pipe][1].split(" ")[0]) + \
                        Fraction(exj_chart[diameter_pipe][1].split(" ")[1])
    free_length_exj = Fraction(exj_chart[diameter_pipe][2].split(" ")[0]) + \
                      Fraction(exj_chart[diameter_pipe][2].split(" ")[1])

    expansion_value = "0/2"  # default zero value
    contraction_value = "0/2"  # default zero value

    if expansion_measurements[0] == 0:
        expansion_value = Fraction(expansion_measurements[1], expansion_measurements[2])
    else:
        expansion_value = Fraction(
            ((expansion_measurements[0] * expansion_measurements[2]) + expansion_measurements[1]),
            expansion_measurements[2])

    if compression_measurements[0] == 0:
        contraction_value = Fraction(compression_measurements[1], compression_measurements[2])
    else:
        contraction_value = Fraction(
            ((compression_measurements[0] * compression_measurements[2]) + compression_measurements[1]),
            compression_measurements[2])

    min_length_exj = Fraction(free_length_exj) - Fraction(max_contraction_exj) + Fraction("1/4")
    max_length_exj = Fraction(free_length_exj) + Fraction(max_expansion_exj) - Fraction("1/4")


    user_setting_input_check = re.findall(r"\d+\'\d+\-\d+\/\d+", user_setting)

    user_setting_length = 0
    if user_setting_input_check:
        print("Input correct!")

        user_setting_length = user_setting.replace('"', '').split("'")
        true_feet = float(user_setting_length[0])
        true_inches = float(user_setting_length[1].split("-")[0]) + float(int(user_setting_length[1].split("-")[1].split("/")[0]) /
                                                                  int(user_setting_length[1].split("-")[1].split("/")[1]))

        user_setting_length = round(true_feet + (true_inches / 12), 4) * 12
        print("User_setting length - ", user_setting_length, "\n")
    else:
        print("Error - input is incorrect")

    status_exp_pipe_vs_exj = 0
    status_con_pipe_vs_exj = 0
    print(Fraction(user_setting_length) - Fraction(expansion_value), Fraction(min_length_exj))
    print(user_setting_length, Fraction(contraction_value), max_length_exj)

    if user_setting_length - Fraction(contraction_value) >= min_length_exj and \
            user_setting_length + Fraction(expansion_value) <= max_length_exj:
        status_con_pipe_vs_exj = 1
        status_exp_pipe_vs_exj = 1
        print("setting is ok")
    else:
        print(Fraction(contraction_value), Fraction(expansion_value))

    print("exp value - ", expansion_value)
    print("con value - ", contraction_value)
    print("user Setting - ", user_setting)

    return





