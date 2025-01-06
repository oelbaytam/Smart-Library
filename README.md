# Smart Library
 A simple macro enabled excel spreadsheet to keep track of your book collection.

## How to Use Smart Library
The purpose of the Smart Library is to keep a record of your manually inputted books that you want, have, and have read. The main benefit of Smart Library is to see what types of books you enjoy reading, it keeps track of the genres and categorizes them into Charts on the charts tab.

The important aspect of using the program is the Database Interactions, the first step for using the library must ALWAYS be importing a database, if kept in the same spot on the computer and not modified outside of the excel application it can be used without reimporting again. Reimporting the Database can be done with a button on the “Data” Sheet or a button on the Library Controls Ribbon. When the Database is imported, it will display all the data on the Bookshelf and Wishlist sheet. 

 
The Bookshelf Sheet after importing a database file.\
The Library Controls Ribbon contains all the controls available, \
1.	Add to Bookshelf – Opens prompt to add item to the bookshelf database.\
2.	Remove from Bookshelf – Removes the Currently selected item from Bookshelf in Sheet 3.\
3.	Add to Wishlist - Opens prompt to add item to the Wishlist database.\
4.	Remove from Wishlist – Removes the Currently selected item from Wishlist in Sheet 4.\
5.	Move to Bookshelf – Moves the selected item from the Wishlist on Sheet 4 to the Bookshelf on Sheet 3.\
6.	Mark as Read – Provides a Date read on selected item from the Bookshelf Table on Sheet 4.\
7.	Mark as Unread – Removes the date on the selected item from the bookshelf item on sheet 4.\
8.	Select a database – Allows you to paste a path to a .mdb file or browse for the file on your system, accepting it refreshes the sheet with its data.\
9.	Generate excel Charts – Allows you to generate one of 3 chart types on Sheet 2 all are Pi Charts, and the chart descriptions are provided beside the buttons.\
10.	Generate Word Report – Generates a word report that creates a document in the location of the workbook named Report.docx, It has all three chart variants and the number of books in the library (the amount owned, and the amount wish listed), the number of books read, and the number of books wish listed.