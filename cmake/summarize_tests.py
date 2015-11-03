import os
import re
import sys

def title_case_to_underscore(name):
	a = re.compile('((?<=[a-z0-9])[A-Z]|(?!^)[A-Z](?=[a-z]))')
	return a.sub(r'_\1', name).lower()

def list_tests(filename):
	f = open(filename)

	in_class = False
	class_name = ''

	for line in f:
		if line.startswith('class '):
			in_class = True
			class_name = line.split(' ')[1][:-2]
			if '(' in class_name:
				class_name = class_name.split('(')[0]
			continue

		if in_class and not (line.startswith('    ') or line.strip() == ''):
			in_class = False
			continue

		if line.lstrip().startswith('def test_'):
			if in_class:
				print('\t' + title_case_to_underscore(class_name) + '__' + line.split()[1].split('(')[0])
			else:
				print('\t' + line.split()[1].split('(')[0])

def find_test_files(dir):
	files = []

	for f in os.listdir(dir):
		if os.path.isdir(os.path.join(dir, f)):
			files.extend(find_test_files(os.path.join(dir, f)))
		elif os.path.isfile(os.path.join(dir, f)) and f.endswith('.py') and f.startswith('test_'):
			files.append(os.path.join(dir, f))

	return files

def main():
	directory = sys.argv[1] if len(sys.argv) > 1 else os.path.dirname(os.path.abspath(__file__))
	test_files = find_test_files(directory)
	for test_file in test_files:
		print(test_file.split('openpyxl/')[1], end='')
		list_tests(test_file)
		print()

if __name__ == '__main__':
	main()