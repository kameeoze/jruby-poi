jruby-poi
=========

This little gem provides an alternative interface to the Apache POI java library, for JRuby. For now the API is targeted at wrapping spreadsheets. We may expand this in the future.

INSTALL
=======

* gem install jruby-poi

USAGE
=====
It's pretty simple really, create a POI::Workbook and access its sheets, rows, and cells. You can read from the spreadsheet, save it (as the filename with which you created or as a new spreadsheet via Workbook#save_as).

    require 'poi'

    # given an Excel spreadsheet whose first sheet (Sheet 1) has this data:
    #    A  B  C  D  E        
    # 1  4     A     2010-01-04
    # 2  3     B     =DATE(YEAR($E$1), MONTH($E$1), A2)
    # 3  2     C     =DATE(YEAR($E$1), MONTH($E$1), A3)
    # 4  1     D     =DATE(YEAR($E$1), MONTH($E$1), A4)
    workbook = POI::Workbook.open('spreadsheet.xlsx')
    sheet = workbook.worksheets["Sheet 1"]
    rows  = sheet.rows
  
    # get a cell's value -- returns the value as its proper type, evaluating formulas if need be
    rows[0][0].value # => 4.0
    rows[0][1].value # => nil
    rows[0][2].value # => 'A'
    rows[0][3].value # => nil
    rows[0][4].value # => 2010-01-04 as a Date instance
    rows[1][4].value # => 2010-01-03 as a Date instance
    rows[2][4].value # => 2010-01-02 as a Date instance
    rows[3][4].value # => 2010-01-01 as a Date instance
    
    # you can access a cell in array style as well... these snippets are all equivalent
    workbook.sheets[0][2][2]          # => 'C'
    workbook[0][2][2]                 # => 'C'
    workbook.sheets['Sheet 1'][2][2]  # => 'C'
    workbook['Sheet 1'][2][2]         # => 'C'

There's a formatted version of this code [here](http://gist.github.com/557607), but Github doesn't allow embedding script tags in Markdown. Go figure!

TODO
====
* fix reading ODS files -- we have a broken spec for this in io_spec.rb
* add APIs for updating cells in a spreadsheet
* create API for non-spreadsheet files

Contributors
============

* [Scott Deming](http://github.com/sdeming)
* [Jason Rogers](http://github.com/jacaetevha)

Note on Patches/Pull Requests
=============================
 
* Fork the project.
* Make your feature addition or bug fix.
* Add tests for it. This is important so I don't break it in a future version unintentionally.
* Commit, do not mess with rakefile, version, or history.
  (if you want to have your own version, that is fine but bump version in a commit by itself I can ignore when I pull)
* Send me a pull request. 

Copyright
=========

Copyright (c) 2010 Scott Deming and others.
See NOTICE and LICENSE for details.

