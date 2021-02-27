import os # operating system
def read_file(filename):
	products = []
	with open (filename, 'r' ,encoding = 'utf-8') as f:
		for line in f:
			if 'name , price' in line :
				continue
			name , price=line.strip().split(',')
			products.append([name, price])		
	return products
def user_input(products):
	while True:
		name =input('please enter your name of the product:    ')
		if name == 'q':
			break
		price = input('Please enter the price :               ')
		products.append([name , price])
	print(products)
	return products
def print_products(products):
	for p in products :
		print('the price of ' ,p[0] ,p[1])
def write_file(filename ,products):
	with open(filename ,'w', encoding = 'utf-8') as f:
		f.write('Product , price\n')
		for p in products :
			f.write(p[0] + ',' +str(p[1]) + '\n')

def main():
	filename = 'products.csv'
	if os.path.isfile(filename) :
		print('yeah ! Found the file')
		products =read_file(filename)
	else:		
		print ('Cannot find it.....')

	
	products = user_input(products)
	print_products(products)
	write_file('products.csv', products)
main()