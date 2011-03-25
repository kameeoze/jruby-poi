#!/bin/sh
#set -x
RUBY_DIR=$(dirname $(which ruby))/..
RUBYGEMS_DIR=${RUBY_DIR}/lib/ruby/gems/jruby/gems

GEM_COLUMNIZE=$(ls -1d $RUBYGEMS_DIR/columnize*/lib | head -1 | /usr/bin/ruby -e 'print File.expand_path($stdin.read)')
GEM_RUBY_DEBUG_BASE=$(ls -1d $RUBYGEMS_DIR/ruby-debug-base-*/lib | head -1 | /usr/bin/ruby -e 'print File.expand_path($stdin.read)')
GEM_RUBY_DEBUG_CLI=$(ls -1d $RUBYGEMS_DIR/ruby-debug-*/cli | head -1 | /usr/bin/ruby -e 'print File.expand_path($stdin.read)')
GEM_SOURCES=$(ls -1d $RUBYGEMS_DIR/sources-*/lib | head -1 | /usr/bin/ruby -e 'print File.expand_path($stdin.read)')

runner="ruby --client \
-I${GEM_COLUMNIZE} \
-I${GEM_RUBY_DEBUG_BASE} \
-I${GEM_RUBY_DEBUG_CLI} \
-I${GEM_SOURCES} \
-rubygems -S"

cmd="spec -c --debugger $*"
#cmd="irb"

$runner $cmd
