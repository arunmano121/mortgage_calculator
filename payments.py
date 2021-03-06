#!/usr/bin/env python3

__authors__ = 'Arun Manohar'
__license__ = '3-clause BSD'
__maintainer__ = 'Arun Manohar'
__email__ = 'arunmano121@outlook.com'

# modules necessary to write out excel files
import xlsxwriter


def write_excel(title, month, interest, principal, out_prin,
                mon_hoa, mon_home_ins, mon_prop_tax, mon_pay,
                home_val, down_pay, loan_amt,
                tot_int, tot_prop_tax, tot_home_ins, tot_hoa, tot_pay,
                int_loan_rat, int_7yr, int_7yr_tot_rat, out_prin_7yr):
    '''
    write out payment schedule into excel sheet

    :param title: name of the excel file
    :type title: str
    :param month: numeric month
    :type month: int
    :param interest: interest component of the monthly payment
    :type interest: float
    :param principal: principal component of the monthly payment
    :type principal: float
    :param out_prin: outstanding principal component owed to bank
    :type out_prin: float
    :param mon_hoa: monthly HOA and Mello-Roos fees
    :type mon_hoa: float
    :param mon_home_ins: monthly home insurance
    :type mon_home_ins: float
    :param mon_prop_tax: monthly property tax
    :type mon_prop_tax: float
    :param mon_pay: total monthly payment
    :type mon_pay: float
    :param home_val: value of home
    :type home_val: float
    :param down_pay: downpayment at time of purchase
    :type down_pay: float
    :param loan_amt: outstanding loan amount
    :type loan_amt: float
    :param tot_int: total interest paid over life of loan
    :type tot_int: float
    :param tot_prop_tax: total property tax over life of loan
    :type tot_prop_tax: float
    :param tot_home_ins: total home insurance over life of loan
    :type tot_home_ins: float
    :param tot_hoa: total HOA and Mello-Roos over life of loan
    :type tot_hoa: float
    :param tot_pay: total payment over life of loan
    :type tot_pay: float
    :param int_loan_rat: interest to loan amount ratio
    :type int_loan_rat: float
    :param int_7yr: interest paid over the first 7 years of loan
    :type int_7yr: float
    :param int_7yr_tot_rat: ratio of interest paid over 7 years to life time
    :type int_7yr_tot_rat: float
    :param out_prin_7yr: outstanding principal remaining after 7 years
    :type out_prin_7yr: float

    :return: None
    '''

    # file name to write out
    workbook = xlsxwriter.Workbook(title)
    worksheet = workbook.add_worksheet()

    # setting up necessary formats
    fmt = workbook.add_format({'bold': True})
    money = workbook.add_format({'num_format': '[$$]#,##0.00'})
    pct = workbook.add_format({'num_format': '0.00%'})

    # write out headers using format
    worksheet.write(0, 0, 'Month', fmt)
    worksheet.write(0, 1, 'Interest', fmt)
    worksheet.write(0, 2, 'Principal', fmt)
    worksheet.write(0, 3, 'HOA', fmt)
    worksheet.write(0, 4, 'Home Ins.', fmt)
    worksheet.write(0, 5, 'Prop. Tax', fmt)
    worksheet.write(0, 6, 'Mon. Payment', fmt)
    worksheet.write(0, 7, 'Out. Principal', fmt)

    # write out individual lines of data
    for i in range(len(month)):
        worksheet.write(i + 2, 0, month[i])
        worksheet.write(i + 2, 1, interest[i], money)
        worksheet.write(i + 2, 2, principal[i], money)
        worksheet.write(i + 2, 3, mon_hoa, money)
        worksheet.write(i + 2, 4, mon_home_ins, money)
        worksheet.write(i + 2, 5, mon_prop_tax, money)
        worksheet.write(i + 2, 6, mon_pay, money)
        worksheet.write(i + 2, 7, out_prin[i], money)

    worksheet.write('J1', 'Home Value', fmt)
    worksheet.write('J2', 'Down Payment', fmt)
    worksheet.write('J3', 'Loan Amt.', fmt)
    worksheet.write('J4', 'Tot. Interest', fmt)
    worksheet.write('J5', 'Tot. Prop. Tax', fmt)
    worksheet.write('J6', 'Tot. Home Ins.', fmt)
    worksheet.write('J7', 'Tot. HOA', fmt)
    worksheet.write('J8', 'Tot. Payment', fmt)
    worksheet.write('J9', 'Int-Loan Ratio', fmt)
    worksheet.write('J11', '7yr Interest', fmt)
    worksheet.write('J12', 'Int 7yr-Total Ratio', fmt)
    worksheet.write('J13', 'Out. Prin. 7yr', fmt)

    worksheet.write('K1', home_val, money)
    worksheet.write('K2', down_pay, money)
    worksheet.write('K3', loan_amt, money)
    worksheet.write('K4', tot_int, money)
    worksheet.write('K5', tot_prop_tax, money)
    worksheet.write('K6', tot_home_ins, money)
    worksheet.write('K7', tot_hoa, money)
    worksheet.write('K8', tot_pay, money)
    worksheet.write('K9', int_loan_rat/100, pct)
    worksheet.write('K11', int_7yr, money)
    worksheet.write('K12', int_7yr_tot_rat/100, pct)
    worksheet.write('K13', out_prin_7yr, money)

    # finished writing - close the workbook
    workbook.close()

    return


def calc_mon_pay(out_prin, months, int_rate):
    '''
    calculate monthly payment including interest and principal

    :param out_prin: outstanding principal amount owed to bank
    :type out_prin: float
    :param months: number of remaining months in loan
    :type months: int
    :param int_rate: fixed interest rate
    :type int_rate: float

    :return: [payment, interest, principal] - list containing total monthly
        payment to bank, interest component in the monthly payment,
        principal component in the monthly payment
    :rtype: list
    '''

    # monthly payment not including home ins and property tax and HOA
    # this only includes the loan amount based payment that is due to bank
    payment = (out_prin * (int_rate / (12 * 100)) /
               (1 - (1 + int_rate / (12 * 100))**(-months)))

    # interest component
    interest = (int_rate/(12 * 100))*out_prin

    # principal component
    principal = payment - interest

    return [payment, interest, principal]


def calc_schedule(loan_amt, years, int_rate):
    '''
    Calculate schedule of payments month over month

    :param loan_amt: outstanding loan amount
    :type loan_amt: float
    :param years: number of years in the loan
    :type months: int
    :param int_rate: fixed interest rate at start of the loan
    :type int_rate: float

    :return: [pay_h, int_h, prin_h, month_h, out_prin_h] -
        list containing monthly total payment to bank, monthly interest
        component to bank, monthly principal component to bank, month of
        payment in numbers 1, 2, 3... etc, outstanding principal after
        current monthly payment
    :rtype: list
    '''

    pay_h = []
    int_h = []
    prin_h = []
    month_h = []
    out_prin_h = []

    # at the very start, the outstanding principal is the loan amount
    out_prin = loan_amt

    # iterate through the life of loan
    for months in range(1, years*12 + 1):
        [payment, interest, principal] = \
            calc_mon_pay(out_prin,
                         years*12 - months + 1, int_rate)

        # outstanding principal reduces every month
        out_prin = out_prin - principal
        # print(months, payment, interest, principal)

        # append the monthly breakdown into the arrays
        pay_h.append(payment)
        int_h.append(interest)
        prin_h.append(principal)
        month_h.append(months)
        out_prin_h.append(out_prin)

        # when the loan is paid off - stop looping, this is relevant when
        # there is additional monthly payments
        if out_prin <= 0:
            break

    return [pay_h, int_h, prin_h, month_h, out_prin_h]


def main():
    '''Calculates the schedule of payments given mortgage parameters.

    The parameters that are involved in calculating the monthly parameters
    are home value, down payment, loan term, interest rate, monthly HOA,
    property tax and home insurance which are roughly based on the property
    tax rates (1.25%). The monthly schedule of payments are output to an
    excel file.
    '''

    # home value
    home_val = float(input('Home value (Million): '))
    # home value in Millions
    home_val *= 1000000
    # downpayment percentage - typically 15% or 20%
    down_pct = float(input('Down-payment (%): '))
    down_pay = home_val * down_pct/100
    # loan term in years
    years = int(input('Loan term (years): '))
    # interest rate
    int_rate = float(input('Interest rate (%): '))
    # loan amount is home value minus down payment
    loan_amt = home_val - down_pay
    # Monthly HOA and Mello-Roos
    mon_hoa = int(input('Monthly HOA and Mello-Roos ($): '))

    # monthly property tax is assumed to be 1.25%
    prop_tax = 1.25
    # monthly property tax
    mon_prop_tax = home_val * prop_tax / 100 / 12
    # typical monthly home insurance
    mon_home_ins = home_val * prop_tax / 100 / 10 / 12

    # calculate the schedule of payments
    print('-'*60)
    print('calculating schedule of payments month over month...')
    print('-'*60)
    [pay_h, int_h, prin_h, month_h, out_prin_h] = \
        calc_schedule(loan_amt, years, int_rate)

    # total monthly payment is sum of payment, hoa, home ins and prop tax
    mon_pay = pay_h[0] + \
        mon_hoa + mon_home_ins + mon_prop_tax
    print('Monthly payment: $%0.2f' % (mon_pay))

    # total interest over the life of loan
    tot_int = sum(int_h)
    print('Total interest payment over %d months: $%0.2f' %
          (years*12, tot_int))

    # total property tax over the life of loan
    tot_prop_tax = mon_prop_tax*12*years
    print('Total taxes over the %d months: $%0.2f' %
          (years*12, tot_prop_tax))

    # total home insurance over the life of loan
    tot_home_ins = mon_home_ins*12*years
    print('Total home insurance over the %d months: $%0.2f' %
          (years*12, tot_home_ins))

    # total HOA over the life of loan
    tot_hoa = mon_hoa*12*years
    print('Total HOA/Mello-Roos over the %d months: $%0.2f' %
          (years*12, tot_hoa))

    # total payment over the life of loan
    tot_pay = down_pay + years*12*mon_pay
    print('Total payment over the %d months: $%0.2f' %
          (years*12, tot_pay))

    # ratio of interest to money borrowed from bank
    int_loan_rat = tot_int/loan_amt*100
    print('Interest-Loan Ratio: %0.2f%%' % (int_loan_rat))

    # interest that is paid over the first 7 years
    int_7yr = sum(int_h[0:7*12])
    print('Interest paid over the first 7 years: $%0.2f' % (int_7yr))

    # Proportion of interest that is paid over the first 7 years
    int_7yr_tot_rat = int_7yr / tot_int * 100
    print('Proportion of total interest paid in first 7 years: %0.2f%%' %
          (int_7yr_tot_rat))

    # Outstanding principal after first 7 years
    out_prin_7yr = out_prin_h[7*12-1]
    print('Outstanding principal after 7 years: $%0.2f' % (out_prin_7yr))

    # write the breakdown into excel file
    print('-'*60)
    print('writing out payment schedule into excel sheet...')
    print('-'*60)

    # title for the excel file
    title = 'schedule_' + \
        str(home_val/1000000) + 'M_' + \
        str(down_pct) + '%dn_' + \
        str(years) + 'yr_' + \
        str(int_rate) + '%int' + \
        '.xlsx'

    # todo: condense necessary items to be written into excel using dict
    write_excel(title, month_h, int_h, prin_h, out_prin_h,
                mon_hoa, mon_home_ins, mon_prop_tax, mon_pay,
                home_val, down_pay, loan_amt,
                tot_int, tot_prop_tax, tot_home_ins, tot_hoa, tot_pay,
                int_loan_rat, int_7yr, int_7yr_tot_rat, out_prin_7yr)


if __name__ == '__main__':
    main()
