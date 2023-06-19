A bill of materials (BOM) is a comprehensive inventory of the raw materials, assemblies, subassemblies, parts and components, as well as the quantities of each needed to manufacture a product. 
In a nutshell, it is the complete list of all the items that are required to build a product. 

It is needed to write a program which can read input Excel file and store the result of calculation to the output Excel file.
The input Excel file contains a Work Sheet with next important columns: "Master-order", "Material reference", "Quantity". 
Need to calculate percentage of similarity of "Material reference" for all "Master-order".

Some features:
  1. Amount of calculated cells can be around 4M.
  2. Parallel FOR implemented.
  3. To avoid form window freezing, TASK class is using.
