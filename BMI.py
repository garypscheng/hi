# BMI
height = input('please enter your height (cm):         ')
weight = input('plase enter your weight with (kg):     ')
height = int(height)
weight = int(weight)
height = height/ 100
BMI =weight/ height / height
if BMI< 18.5:
	print('Your BMI:','underweight')
elif BMI >= 18.5 and BMI <= 24.9:
	print('Your BMI :','Nomral weight')
elif BMI >= 25 and BMI <= 29.9:
	print ('Your BMI :','Overweight')
else:
	print('Your BMI :','Obesity')
