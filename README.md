This is the homepage for paper "Automated Refactoring of Nested-IF Formulae in Spreadsheets". 
----------------------------------------------------------------------------------------------------

Abstract:

Spreadsheets are the most popular end-user programming software, where formulae act like programs and also have smells. One well recognized common smell of spreadsheet formulae is nest-if expressions, which have low readability and high cognitive cost for users, and are error-prone during reuse or maintenance. However, end users usually lack essential programming language knowledge and skills to tackle or even realize this problem, yet no effective and automated approach is currently available to provide support for end users. 

This paper proposes the first automated approach to systematically refactoring nest-if formulae. The general idea is two-fold. First, we detect and remove logic redundancy based on the AST of a formulae. Second, we identify higher-level semantics that have been represented with fragmented and scattered syntax, and reassemble the syntax using concise built-in functions. A comprehensive evaluation has been conducted against two large-scale real-world spreadsheet corpora. The results with over 80,000 spreadsheets and over 28 million nest-if formulae. reveal that the approach is able to relieve the smell of over 90% of nest-if formulae. A survey involving 49 participants also indicates that for most cases the participants prefer the refactored formulae, and agree on that such automated refactoring approach is necessary and helpful.

----------------------------------------------------------------------------------------------------

Package NestIF-PythonCode contains the code for:
1. get refactored formulae from the original nest-if formulae (artifact1_get_refactored_formulae.py).
2. analyze the refactored results (artifact2_analyze_refactored_formulae.py).


Link https://drive.google.com/open?id=1003s1HfvFBcBbdt3IQ0s29XwSHrb9icE contains the data of:
1. the original nest-if formulae extracted from the Enron spreadsheet corpus (). 
2. the refactored formulae.

----------------------------------------------------------------------------------------------------

Please follow these steps to run experiments:

Step1. Download the code and data. Record the filepath of package "refactor-data".

Step2. In the main class of artifact1_get_refactored_formulae.py, change the filepath into the path of folder "refactor-data".

Step3. Run artifact1_get_refactored_formulae.py to get the refactored results.

Step4. In the main class of artifact2_analyze_refactored_formulae.py, change the filepath into the path of folder "refactor-data".

Step5. Run artifact2_analyze_refactored_formulae.py to get the analyzed results.
