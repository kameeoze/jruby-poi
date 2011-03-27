#!/bin/sh
#set -x
RUBY_DIR=$(dirname $(which ruby))/..
if [[ ${RUBY_DIR} == *1.6.* ]]
then
	RUBYGEMS_DIR=${RUBY_DIR}/lib/ruby/gems/jruby/gems
else
	RUBYGEMS_DIR=${RUBY_DIR}/lib/ruby/gems/1.8/gems
fi

GEM_COLUMNIZE=$(ls -1d $RUBYGEMS_DIR/columnize*/lib | head -1 | /usr/bin/ruby -e 'print File.expand_path($stdin.read)')
GEM_RUBY_DEBUG_BASE=$(ls -1d $RUBYGEMS_DIR/ruby-debug-base-*/lib | head -1 | /usr/bin/ruby -e 'print File.expand_path($stdin.read)')
GEM_RUBY_DEBUG_CLI=$(ls -1d $RUBYGEMS_DIR/ruby-debug-*/cli | head -1 | /usr/bin/ruby -e 'print File.expand_path($stdin.read)')
GEM_SOURCES=$(ls -1d $RUBYGEMS_DIR/sources-*/lib | head -1 | /usr/bin/ruby -e 'print File.expand_path($stdin.read)')

echo "RUBYGEMS_DIR:        ${RUBYGEMS_DIR}"
echo "GEM_SOURCES:         ${GEM_SOURCES}"
echo "GEM_COLUMNIZE:       ${GEM_COLUMNIZE}"
echo "GEM_RUBY_DEBUG_CLI:  ${GEM_RUBY_DEBUG_CLI}"
echo "GEM_RUBY_DEBUG_BASE: ${GEM_RUBY_DEBUG_BASE}"

runner="ruby --client \
-I${GEM_COLUMNIZE} \
-I${GEM_RUBY_DEBUG_BASE} \
-I${GEM_RUBY_DEBUG_CLI} \
-I${GEM_SOURCES} \
-rubygems -S"

cmd="bundle exec rdebug rspec -c $*"
#cmd="irb"

$runner $cmd
