# xsl2vCard
convert xsl file to vCards. One for each contact in your xsl file. Handy when migrating mobile platforms.


Leveraging the Apache's librarys, this tool will help you export vCards from xls file rows.

e.x. suppose you have a xls file like so:
<pre>
#######################################
#First Name # Last Name # ph.Number   #
#Jean       # Doeveen   # 12-345-567  #
#John       # Doe       # 2345678901  #
#######################################
</pre>
and you want to import the contacts to some cloud service that supports .vCard file format.
Change the TOTALROWS, take care of the paths and you are ready to go!

This tool will allow easy export of any number of lines of contacts, each to a seperate .vCard file for ease of manipulation.
In case of null values, a series of try-catch are implemented to prevent runtime exceptions.
Athough there are more to be done in order for this chunk to be fully stable, it is a prototype i used to make my life easier.
