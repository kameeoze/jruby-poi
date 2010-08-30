jruby-poi
=========

This little gem provides an alternative interface to the Apache POI java library, for JRuby. For now the API is targeted at wrapping spreadsheets. We may expand this in the future.

INSTALL
=======

* gem install jruby-poi

USAGE
=====
It's pretty simple really, create a POI::Workbook and access its sheets, rows, cells -- read from it, write to it.
<script src="http://gist.github.com/557607.js?file=gistfile1.rb"></script>
TODO
====
* fix reading ODS files -- we have a broken spec for this in io_spec.rb
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

