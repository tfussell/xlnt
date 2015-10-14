SCRIPT_DIR="$( cd "$( dirname "${BASH_SOURCE[0]}" )" && pwd )"

# use included genie if it isn't found in PATH
GENIE_BIN=$(which genie)
if [ -z "$GENIE_BIN" -o ! -x "$GENIE_BIN" ]; then
    GENIE_BIN=$(pwd)"/genie-osx"
fi

# default
ACTION="xcode4"
if [ ! -z "$1" ]; then
    ACTION="$1"
fi

if [[ "$ACTION" = "clean" ]]; then
    rm -rf xcode4
    rm -rf gmake
else
    $GENIE_BIN $ACTION > /dev/null
    if [[ "$ACTION" = "xcode4" ]]; then
	cd xcode4
	xcodebuild -workspace xlnt.xcworkspace -scheme xlnt.test
    fi
fi
