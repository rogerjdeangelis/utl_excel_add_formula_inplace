    ```  Updating an excel tabl with formulas in place  ```
    ```    ```
    ```  R and python are more flexible, however is is very clean in SAS.  ```
    ```  I have posted more sophisticated updates with R.  ```
    ```  The SAS method is not a general method, but should work for many cases  ```
    ```  of updates.  ```
    ```    ```
    ```    WORKING CODE  ```
    ```    ```
    ```        * create update dataset  ```
    ```        data class;  ```
    ```          set sashelp.class(obs=3);  ```
    ```          age = 10 + _n_;  ```
    ```          height=10*_N_ + 80;  ```
    ```    ```
    ```        * modify template with new data and calculate BMI;  ```
    ```        data xel.have;  ```
    ```          modify xel.have class;  ```
    ```          by name;  ```
    ```        run;quit;  ```
    ```    ```
    ```  see  ```
    ```  https://communities.sas.com/t5/Base-SAS-Programming/EXCEL-EXPORT/m-p/369709  ```
    ```    ```
    ```  CAVEATS  ```
    ```    ```
    ```    1. You need a primary key, name works for sashelp.class.  ```
    ```       It is generally good to have a single or compound primary key.  ```
    ```    ```
    ```    2. You only have to do the following steps once on the template excel sheet.  ```
    ```         a. I suggest you programatically make a copy of the template  ```
    ```            since you will may want to use it for something else.  ```
    ```         b If the formulas in the template show and not the values do the following.  ```
    ```            Highlight the column with the formulas, go to find and replace  ```
    ```            find '=' and replace with '-'. This will force the formulas  ```
    ```            to calculate.  ```
    ```    ```
    ```    ```
    ```  HAVE (Template with Sample data and I want to update age, sex, height and weight)  ```
    ```  =================================================================================  ```
    ```    ```
    ```  d:/xls/formulas.xlsx  ```
    ```    ```
    ```      +----------------------------------------+------------------------+  ```
    ```      |  A    |  B    |  C    |  D    |   E    |        F               |  ```
    ```      +----------------------------------------+------------------------+  ```
    ```  1   |NAME   |AGE    |SEX    |HEIGHT |WEIGHT  |        BMI             |  ```
    ```      |-------+-------+-------+-------+--------|------------------------|  ```
    ```  2   |Alex   |14     |M      |69     |112.5   |=E2/(2.2*(D2/39.37)^2)  |  ```
    ```      |-------+-------+-------+-------+--------+------------------------+  ```
    ```  3   |Alice  |13     |F      |56.5   |84      |=E3/(2.2*(D3/39.37)^2)  |  ```
    ```      |-------+-------+-------+-------+--------+------------------------+  ```
    ```  4   |Barbara|13     |F      |65.3   |98      |=E4/(2.2*(D4/39.37)^2)  |  ```
    ```      |-------+-------+-------+-------+--------+------------------------+  ```
    ```  ...  ```
    ```    ```
    ```  [HAVE]  ```
    ```    ```
    ```  WANT (update age and height)  ```
    ```  =============================  ```
    ```    ```
    ```  Note AGE and HEIGHT have changed and new BMI has ben calculated  ```
    ```    ```
    ```  d:/xls/formulas.xlsx  ```
    ```    ```
    ```      +----------------------------------------+--------------+  ```
    ```      |  A    |  B    |  C    |  D    |   E    |      F       |  ```
    ```      +----------------------------------------+--------------+  ```
    ```  1   |NAME   |AGE    |SEX    |HEIGHT |WEIGHT  |     BMI      |  ```
    ```      |-------+-------+-------+-------+--------|--------------|  ```
    ```  2   |Alex   |11     |M      |90     |112.5   |9.785333965   |  ```
    ```      |-------+-------+-------+-------+--------+--------------+  ```
    ```  3   |Alice  |12     |F      |100    |84      |5.918169982   |  ```
    ```      |-------+-------+-------+-------+--------+--------------+  ```
    ```  4   |Barbara|13     |F      |110    |98      |5.7062245     |  ```
    ```      |-------+-------+-------+-------+--------+--------------+  ```
    ```  ...  ```
    ```    ```
    ```  [HAVE]  ```
    ```    ```
    ```  *                _          _                       _       _  ```
    ```   _ __ ___   __ _| | _____  | |_ ___ _ __ ___  _ __ | | __ _| |_ ___  ```
    ```  | '_ ` _ \ / _` | |/ / _ \ | __/ _ \ '_ ` _ \| '_ \| |/ _` | __/ _ \  ```
    ```  | | | | | | (_| |   <  __/ | ||  __/ | | | | | |_) | | (_| | ||  __/  ```
    ```  |_| |_| |_|\__,_|_|\_\___|  \__\___|_| |_| |_| .__/|_|\__,_|\__\___|  ```
    ```                                               |_|  ```
    ```  ;  ```
    ```    ```
    ```  %utlfkil(d:/xls/formulas.xlsx);  ```
    ```  libname xel "d:/xls/formulas.xlsx" scan_text=no;  ```
    ```  data xel.have;  ```
    ```     set sashelp.class(obs=3);  ```
    ```     bmi=cats('=E',put(_n_+1,3.),'/(2.2*(D',put(_n_+1,3.),'/39.37)^2)');  ```
    ```     putlog bmi=;  ```
    ```  run;quit;  ```
    ```  libname xel clear;  ```
    ```    ```
    ```  /*  ```
    ```     Open up excel template. Only need to do this once on template  ```
    ```     Highlight the column with the formulas, go to find and replace  ```
    ```     find '=' and replace with '-'. This will force the formulas  ```
    ```     to calculate.  ```
    ```  */  ```
    ```    ```
    ```  *          _       _   _  ```
    ```   ___  ___ | |_   _| |_(_) ___  _ __  ```
    ```  / __|/ _ \| | | | | __| |/ _ \| '_ \  ```
    ```  \__ \ (_) | | |_| | |_| | (_) | | | |  ```
    ```  |___/\___/|_|\__,_|\__|_|\___/|_| |_|  ```
    ```    ```
    ```  ;  ```
    ```    ```
    ```  libname xel "d:/xls/formulas.xlsx" scan_text=no;  ```
    ```  data class;  ```
    ```    set sashelp.class(obs=3);  ```
    ```    age = 10 + _n_;  ```
    ```    height=10*_N_ + 80;  ```
    ```  run;quit;  ```
    ```    ```
    ```  data xel.have;  ```
    ```    modify xel.have class;  ```
    ```    by name;  ```
    ```  run;quit;  ```
    ```  libname xel clear;  ```
    ```    ```
    ```    ```
