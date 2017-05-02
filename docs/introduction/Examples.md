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

### Simple - storing a spread sheet in a 2 dimensional C++ Vector for further processing
Loading a spread sheet into a Vector provides oppourtunities for you to perform high performance processing. There will be more examples on performing fast look-ups, merging data, performing deduplication and more. For now, let's just learn how to get the spread sheet loaded into memory.
```
#include <iostream>
#include <xlnt/xlnt.hpp>
#include <vector>

using namespace std;
using namespace xlnt;


int main()
{
    workbook wb;
    wb.load("/home/timothymccallum/test.xlsx");
    auto ws = wb.active_sheet();
    clog << "Processing spread sheet" << endl;
    clog << "Creating a single vector which stores the whole spread sheet" << endl;
    vector< vector<string> > theWholeSpreadSheet;
	for (auto row : ws.rows(false)) 
	    { 
                clog << "Creating a fresh vector for just this row in the spread sheet" << endl;
		vector<string> aSingleRow;
		for (auto cell : row) 
		    { 
			clog << "Adding this cell to the row" << endl;
			aSingleRow.push_back(cell.to_string());
		    }
		clog << "Adding this entire row to the vector which stores the whole spread sheet" << endl;
		theWholeSpreadSheet.push_back(aSingleRow);
	    }
    clog << "Processing complete" << endl;
    clog << "Reading the vector and printing output to the screen" << endl;
    for (int rowInt = 0; rowInt < theWholeSpreadSheet.size(); rowInt++)
	{
	    for (int colInt = 0; colInt < theWholeSpreadSheet.at(rowInt).size(); colInt++)
		{
		    cout << theWholeSpreadSheet.at(rowInt).at(colInt) << endl;
		}
	}
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
Creating a single vector which stores the whole spread sheet
Creating a fresh vector for just this row in the spread sheet
Adding this cell to the row
Adding this cell to the row
Adding this cell to the row
Adding this entire row to the vector which stores the whole spread sheet
Creating a fresh vector for just this row in the spread sheet
Adding this cell to the row
Adding this cell to the row
Adding this cell to the row
Adding this entire row to the vector which stores the whole spread sheet
Processing complete
Reading the vector and printing output to the screen
This is cell A1.
This is cell B1
… and this is cell C1
We are now on the second row at cell A2
B2
C2
```
You will have noticed that this process is very fast. If you type the "time" as shown below, you can measure just how fast loading and retrieving your spread sheet is, using xlnt; In this case only a fraction of a second. More on this later.
```
time ./process 
...
real	0m0.044s
```
### Simple - writing values to a new xlsx spread sheet.
