# DocumentContentsScan
A data extraction script; parses the contents of all documents (Word, Excel and PDFs) in a directory and outputs them into Excel specific spreadsheet.

This was done as a learning project with strong utility but also with some heavy help from ChatGPT4. Enough was written by GPT that I can't claim credit or lisence. As far as I can tell, this should work on any computer. In the code, simply change the dir_paths = [""] to the directory needing to be scanned and it'll go to work. This is an effective script, but it is also not super fast, especially with a large repository of documents. In my usage, about 50,000 documents takes about 8-9 hours to parse.
