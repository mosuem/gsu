# gsu

For Schwab:

Follow the instructions at

https://screenshot.googleplex.com/8rYj3UYcgwKnW8r

Go to 

https://gsu-calculator.web.app

Click the button and upload the JSON.

It should download a XLSX. Retrieve the value under TOTAL.

Profit!


## CLI usage

Run
```bash
cd gsu_cli
dart run bin/gsu_cli.dart --inputFile EquityAwardsCenter_Transactions.json --templateFile ../gsu_app/assets/gsu_template.xlsx --outputFolder .
```