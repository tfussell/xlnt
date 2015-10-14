cd genie
export GENIE_BIN=$(which geni)
if [ ! -x $GENIE_BIN ]; then
    export GENIE_BIN="genie-osx"
    echo "a"
fi
echo $GENIE_BIN
