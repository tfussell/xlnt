## Examples
### Simple - reading from an existing xlsx spread sheet.
The following C plus plus code will read the values from an xlsx file and print the string values to the screen. This is a very simple example to get you started.

```
#include <iostream>
#include <xlnt/xlnt.hpp>

using namespace std;
using namespace xlnt;

int main()
{
	workbook wb;
	wb.load("/home/timothymccallum/test.xlsx");
    auto ws = wb.active_sheet();
    clog << "Processing spread sheet" << endl;
	for (auto row : ws.rows(false)) 
		{ 
			for (auto cell : row) 
			{ 
			    clog << cell.to_string() << endl;
			}
		}
	clog << "Processing complete" << endl;
	return 0;
}
```
Save the contents of the above file 
```
/home/timothymccallum/process.cpp
```
Compile by typing the following command
```
g++ -std=c++14 -Ixlnt/include -Lxlnt/lib -lxlnt process.cpp -o process
```
Excecute by typing the following command
```
./process
```
The output of the program, in my case, is as follows
```
Processing spread sheet
This is cell A1.
This is cell B1
… and this is cell C1
We are now on the second row at cell A2
B2
C2
Processing complete
```
As you can see the process.cpp file simply walks through the spread sheet values row by row and column by column (A1, B1, C1, A2, B2, C2 and so on).

### Simple - writing values to a new xlsx spread sheet.
