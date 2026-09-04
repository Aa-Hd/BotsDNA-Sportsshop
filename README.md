This bot opens https://botsdna.com/sportshop/, 
downloads the School Email Database,
scrapes the latest product list (name/code/unit price), 
reads the list of schools from the Schools page, 
then for each school: 
  searches for that school on the database to reveal its discount % 
  and get school's mail address, 
  builds a "Price Quotation" For Each School in a New document (same layout as SportsTemplet.pdf) with offer prices computed from the discount, 
  and emails it to that school's address. 
Sends a summary email at the end to admin, 
and a failure email if the run breaks.
