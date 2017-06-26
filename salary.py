# dann254

#import xlrd and xlwt
import xlrd, xlwt

#open the excel document to be modified and specify the sheet number.
workbook_a = xlrd.open_workbook('flowers.xlsx')
flower_database = workbook_a.sheet_by_index(0)

#check the length of the document using the 4th column of the document.
flower_col_len = len(flower_database.col_values(3))

#create an sheet to save the results on.
resultbook = xlwt.Workbook()
ws = resultbook.add_sheet('result sheet', cell_overwrite_ok=True)

#define a function that calculates raw salary.
def raw_salary(flowers):
        salary = flowers*10
        return salary

#define a function that checks if the badly cut flowers are beyond 100 and calculate the deductions.
def bad_cuts(bad_flower):
        if bad_flower > 100:
                bad_count = bad_flower - 100
                dd = bad_count*5
                return dd
        else:
                return 0

#define the function that calculates bonus salary for more than 5000 flowers cut.
def five_K_bonus(flowers):
        if flowers > 5000:
                bonus_count = flowers - 5000
                bns = bonus_count*5
                return bns
        else:
                return 0

#define a function that calculates bonus for less than 10 badly cut flowers.
def less_bad_bonus(bad_flower, raw_sal):
        if bad_flower < 10:
                accuracy_bns= 0.03*raw_sal
                return accuracy_bns
        else:
                return 0

#create styles for the entire document.     
style_string = "font: bold on; borders: bottom thin, right thin, left thin, top thin; pattern: pattern solid, fore_colour aqua; alignment: horiz centre, vert centre, wrap on;"
style = xlwt.easyxf(style_string)

style_string2 = "borders: bottom thin, right thin, left thin, top thin;"
style2 = xlwt.easyxf(style_string2)

style_string3 = "font: italic on,bold on, height 330, color red; pattern: pattern solid, fore_colour aqua;"
style3 = xlwt.easyxf(style_string3)

style_string4 = "font: italic on; pattern: pattern solid, fore_colour aqua;"
style4 = xlwt.easyxf(style_string4)

style_string5 = "pattern: pattern solid, fore_colour aqua;"
style5 = xlwt.easyxf(style_string5)

style_string6 = "font: bold on; pattern: pattern solid, fore_colour aqua;"
style6 = xlwt.easyxf(style_string6)

#write the styles on the document titles and save the titles to the result sheets.
ws.write(1, 1, "Flowers.ke", style3)
ws.write(2, 1, "for the finest flowers in kenya", style4)
ws.write(3, 1, "June 2017", style5)
ws.write(4, 1, "Employee Flower records",style6)

ws.write(1, 2, "",style5)
ws.write(2, 2, "",style5)
ws.write(3, 2, "",style5)
ws.write(4, 2, "",style5)

ws.write(1, 3, "",style5)
ws.write(2, 3, "",style5)
ws.write(3, 3, "",style5)
ws.write(4, 3, "",style5)

#write syles on the collumn headings and save to the result sheet
ws.write(8, 1, "Id", style)
ws.write(8, 2, "Name", style)
ws.write(8, 3, "Flowers cut", style)
ws.write(8, 4, "Badly cut flowers", style)
ws.write(8, 5, "Raw salary (ksh)", style)
ws.write(8, 6, "Deduction (ksh)", style)
ws.write(8, 7, "Bonus (ksh)", style)
ws.write(8, 8, "Final salary (ksh)", style)

#initialize totals
flower_total = 0
bad_total = 0
raw_total = 0
deduction_total = 0
bonus_total = 0
salary_total = 0

#reduce the column lenght by 1 to get rid of the totals row before looping through it.
f_len = flower_col_len-1

#create loop to do the calculations for each employee
for i in range(9,f_len):
        #read the data from the loaded document
        employee_id = str(flower_database.cell(i, 1).value)
        employee_name = str(flower_database.cell(i, 2).value)
        flowers_cut = int(flower_database.cell(i, 3).value)
        badly_cut_flowers = int(flower_database.cell(i, 4).value)
        
        flowers = int(flowers_cut)
        bad_flower = int(badly_cut_flowers)

        #pass the read data to the functions above and save the returned result.
        raw_sal = raw_salary(flowers)
        deductibles = bad_cuts(bad_flower)
        five_bonus = five_K_bonus(flowers)
        accuracy_bonus = less_bad_bonus(bad_flower,raw_sal)

        #calculate the bonus totals and salary totals.
        bonus = five_bonus+accuracy_bonus
        final_salary = (raw_sal + bonus) - deductibles

        # save the data to the result sheet and add styles.
        ws.write(i,1, employee_id, style2)
        ws.write(i,2, employee_name, style2)
        ws.write(i,3, flowers_cut, style2)
        ws.write(i,4, badly_cut_flowers, style2)
        ws.write(i,5, raw_sal, style2)
        ws.write(i,6, deductibles, style2)
        ws.write(i,7, bonus, style2)
        ws.write(i,8, final_salary, style2)

        #calculate totals for each collumn
        flower_total = flower_total + flowers_cut
        bad_total = bad_total + badly_cut_flowers
        raw_total = raw_total + raw_sal
        deduction_total = deduction_total + deductibles
        bonus_total = bonus_total + bonus
        salary_total = salary_total + final_salary
        
n = flower_col_len - 1

#save the totals to the result sheet.
ws.write(n,1, "Totals", style)
ws.write(n,2, "", style)
ws.write(n,3, flower_total, style)
ws.write(n,4, bad_total, style)
ws.write(n,5, raw_total, style)
ws.write(n,6, deduction_total, style)
ws.write(n,7, bonus_total, style)
ws.write(n,8, salary_total, style)     


#save the resultsheet in a workbook.
resultbook.save("Salaries.xls")
print "Calculations complete.. check salaries.xls for the result"



        




