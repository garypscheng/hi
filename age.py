driving= input("Have your driven care before?      ")
if driving != 'Yes' and driving != 'No':
	print('Merely Enter Yes/No')
	raise SystemExit
age = input ('Please enter your age:               ')
age = int(age)
if driving == 'Yes':
	if age >= 18:
		print('You passed the test')
	else:
  		print('You failed the test')
elif driving == 'No':
	if age >= 18:
		print( 'You can join the driving test to get passbook')
	else: 
		print('Maybe, you need to wait for serval years')
