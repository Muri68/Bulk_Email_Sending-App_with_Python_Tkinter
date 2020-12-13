import pandas as pd #pip install pandas

data = pd.read_excel("withemail.xlsx")
#print(data)

if 'Email' in data.columns:
	emails = list(data['Email'])
	c = []
	for i in emails:
		# print(i)
		if pd.isnull(i)==False:
			# print(i)
			c.append(i)

	emails = c
	print(emails)
else:
	print("Not Exist")