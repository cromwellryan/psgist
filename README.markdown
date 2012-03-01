PSGist
====================

Provides the `gist` Powershell function to post code to http://gist.github.com.

Usage
--------------------
	`gist [[-File] <FileInfo>] [[-Description] <String>] [[-Username] <String>]` 

Example:

	"Hello World!" | out-file "Greeting.txt"
	gist -File "Greeting.txt" -Desription "My First PsGist" -Username cromwellryan

