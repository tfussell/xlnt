SCRIPT_DIR="$( cd "$( dirname "${BASH_SOURCE[0]}" )" && pwd )"

# use included genie if it isn't found in PATH
GENIE_BIN=$(which genie)
if [ -z "$GENIE_BIN" -o ! -x "$GENIE_BIN" ]; then
    GENIE_BIN=$(pwd)"/genie-linux"
fi

# default
ACTION="gmake"
if [ ! -z "$1" ]; then
    ACTION="$1"
fi

if [[ "$ACTION" = "clean" ]]; then
    rm -rf gmake
else
    $GENIE_BIN $ACTION > /dev/null
    if [[ "$ACTION" = "gmake" ]]; then
	cd gmake
	make
    fi
fi
