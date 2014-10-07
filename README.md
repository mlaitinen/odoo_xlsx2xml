odoo_xlsx2xml
=============

A small script for converting a straight xlsx file to an odoo xml file.

Usage: ```odoo_xlsx2xml.py <xlsx file> <record model> [<output.xml>]```

Example: ```odoo_xlsx2xml.py accounts.xlsx account.account.template chart_of_accounts.xml```

The xlsx file must meet the following requirements:
 - contains only one worksheet
 - field names are located on the first row of the sheet
 - contains a field called ```id```
 - the data starts from row 2

A column can contain attribute values instead of regular text content. These columns must have
question mark as a delimiter and the attribute name after that. See example below (parent?ref).
<br />
Converts this:

record model: ```my.module.label```

id    | name    | parent?ref
:---- | :-------| :---------
a0    | Record 1|
a1    | Record 2| a0

<br />
into this:
```
<openerp>
    <data noupdate="1">
        <record model="my.module.label" id="a0">
            <field name="name">Record 1</field>
        </record>
        <record model="my.module.label" id="a1">
            <field name="name">Record 2</field>
            <field name="parent" ref="a0" />
        </record>
    </data>
</openerp>
```
