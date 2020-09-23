SQL Toolz ReadMe.

SQL Toolz is an interactive query tool for Adaptive Server Anywhere
and SQL Anywhere.  It has been tested with SQLA 5.5,
ASA 6, 7 and a little with 8.

If this is the patch version of SQL Toolz, simply copy the
included files to the program directory.

SQL Toolz requires certain components which now come with most
versions of Windows.  In order to make the download size smaller,
these are no longer included.  They can however be downloaded from
the Microsoft website.  If you have trouble running SQL Toolz,
please ensure that you have these components installed.

CMAX OCX:  This is the ocx for the editor.  You can download it from:
http://www.compiled.org/files/115/259/cmax20.ocx

RDO:  SQL Toolz uses RDO for its data access layer.

Visual Basic 6 Runtime:  This is the runtime for VB6.

ODBC:  SQL Toolz has been tested with ODBC 3.5.

OLE32:  Updated OLE support as well as service packs can be
obtained for various operating systems from Microsoft.

DCOM Support:  DCOM comes with Windows 98+, Windows NT+, and some
versions of Windows 95.  Older versions of Windows 95 do not
include DCOM support.  This can be obtained from Microsoft.

There is an issue with DDL generation and the Float datatype. The
Float datatype will accept a precision when created, but the DDL
generation does not include the precision, even if one was
specified. This is under investigation.

See revision.txt for revision history.

All names are trademarks and/or copyright of their respective
owners.

SQL Toolz Copyright 1998-2002 Larry W. Stavinoha
http://www.stavinoha.com/pub/sqltoolz.htm
