# PDM_SW_FB_API_V2
Similar to previous PDM / SW puller, but pulls exclusively from the PDM and creates a *.csv that FishBowl can import directly and create a kit with.

Factoral Dependencies:

The Assembly must be named "PROD-####.SLDASM".

The PROD in question must be created by CWI standards, meaning it has a folder
containing all other standard folders (i.e. CNC, Design Reference, etc.)

Assembly and Multibody part must be released.

Any part other than a SWOOD material must be in a library 
with a naming convention that ends with the FishBowl part number.

The part number of each part must exist. If the program produces a "N/A", it doesn't exist in FishBowl.
