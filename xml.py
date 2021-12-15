import xlsxwriter
workbook = xlsxwriter.Workbook('hello.xlsx')
 

worksheet = workbook.add_worksheet()
 
worksheet.write('A1', 'Informatii Generale')
worksheet.write('A2', 'PN')
worksheet.write('B2', 'GPL')
worksheet.write('C2', 'Discount')
worksheet.write('D2', 'Quantity')

worksheet.write('A3', 'MX67-HW')
worksheet.write('A4', 'LIC-MX67-ENT-3YR')

worksheet.write('A7', 'Informatii specifice HQ')
worksheet.write('A8', 'Access Points')
worksheet.write('A9', 'MR46-HW')
worksheet.write('A10', 'LIC-MR-ENT-3YR')

worksheet.write('A11', 'Switches')
worksheet.write('A12', 'MS210-48FP-HW')
worksheet.write('A13', 'LIC-MS-ENT-3YR')

worksheet.write('A14', 'MX virtual pentru cloud')
worksheet.write('A15', 'LIC-VMX-S-ENT-3Y')

worksheet.write('A16', 'Camera supraveghere')
worksheet.write('A17', 'MV22-HW')
worksheet.write('A18', 'LIC-MV-3YR')

worksheet.write('A19', 'Senzori')
worksheet.write('A20', 'MT10-HW')
worksheet.write('A21', 'LIC-MT-3Y')


worksheet.write('A23', 'Informatii specifice Branch')
worksheet.write('A24', 'Access Points')
worksheet.write('A25', 'MR46-HW')
worksheet.write('A26', 'LIC-MR-ENT-3YR')

worksheet.write('A27', 'Switches')
worksheet.write('A28', 'MS210-48FP-HW')
worksheet.write('A29', 'LIC-MS-ENT-3YR')

worksheet.write('A30', 'MX virtual pentru cloud')
worksheet.write('A31', 'LIC-VMX-S-ENT-3Y')

worksheet.write('A32', 'Camera supraveghere')
worksheet.write('A33', 'MV22-HW')
worksheet.write('A34', 'LIC-MV-3YR')

worksheet.write('A35', 'Senzori')
worksheet.write('A36', 'MT10-HW')
worksheet.write('A37', 'LIC-MT-3Y')

worksheet.write('A39', 'Informatii specifice depozit')

worksheet.write('A40', 'Access Points Interior')
worksheet.write('A41', 'MR46-HW')
worksheet.write('A42', 'LIC-MR-ENT-3YR')

worksheet.write('A43', 'Access Points Exterior')
worksheet.write('A44', 'MR84-HW')
worksheet.write('A45', 'LIC-MR-ENT-3YR')

worksheet.write('A46', 'Switches')
worksheet.write('A47', 'MS210-48FP-HW')
worksheet.write('A48', 'LIC-MS-ENT-3YR')

worksheet.write('A49', 'MV22-HW')
worksheet.write('A50', 'LIC-MV-3YR')


worksheet.write('A51', 'MT10-HW')
worksheet.write('A52', 'LIC-MT-3Y')

worksheet.write('A53', '20')
worksheet.write('A54', '30')

worksheet.write('A55', '=SUM(A54 + A53)')



 
workbook.close()