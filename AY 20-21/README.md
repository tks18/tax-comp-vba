### Income Tax Computation in a Single Function
##### Done by [Shan.tk](https://github.com/tks18)

***Applicable for AY 20-21, 21-22***

**A Single Function that Calculates the Income Tax as per Income Tax Act with Multiple Options, Including Surcharge Applicable for Multiple Slabs and Education Cess.**

#### How to Use ?

  * Navigate to Tax Computation - AY 20-21.bas file in my Github Repo - [Click Here](https://github.com/tks18/tax-comp-vba/blob/main/AY%2020-21/Tax%20Computation%20-%20AY%2020-21.bas)

  * Click on Raw Button on top of the File.

  * Copy its Contents

  * Open Excel

  * Press `Alt + F11` for Visual Basics Tab

  * Right Click on the Worksheet Name in left Tab and Inset New Module

  * Paste the Contents of my File in the Module

  * Again Press `Alt + F11` for Cell View

  * Now You can Access the Function in any Cell by `+INDTAX()`

#### Available Options:

  **SYNTAX:`+INDTAX(INCOME_AMOUNT, DISPLAY_OPTIONS(Optional))`**

  1. *INCOME_AMOUNT*:
    * Here You can Enter the Amount of Income You Have Earned. Nothing Complicated hehe xD

    * Example: `+INDTAX(11111111)`

  2. *DISPLAY_OPTIONS___(Optional)___*:
    * Here You can Enter Following Options based on Your Choice:
      * "s1" - For 1st Slab Amount
      * "s2" - For 2nd Slab Amount
      * "s3" - For 3rd Slab Amount
      * "surch" - For Surcharge Amount if Any
      * "cess" - For Cess Amount
      * "noround" - Original Tax will be Rounded as per Income Tax Act, this Option will give you a Tax that is not Rounded as per the Act.
