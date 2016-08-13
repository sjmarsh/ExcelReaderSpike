# Excel Reader Spike #
Needed to read an Excel spreadsheet into a generic List<T> with the ability to map a column to a property without relying on the column header in the spreadsheet. Instead I wanted to create a mapping between the property name of the type and the spreadsheet column identifier (ie. Column A, Column B, etc.).  Many existing solutions use the use the header row of the spreadsheet however we cannot alway rely on it being correct (may have been tampered with).  Also I am using the EPPlus package to read the spreadsheet as it does not rely on any additional Office SDK installations. 
