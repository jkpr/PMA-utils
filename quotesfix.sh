#!/usr/bin/env bash

stringFile=xl/sharedStrings.xml

function usage {
    echo "Usage: $0 [-h] XLSXFILE [XLSXFILE ...]"
}

function showHelpAndExit {
    usage
    echo ""
    echo "Convert all unicode quote characters to ASCII apostrophe or "
    echo "quotation mark in Excel workbook. Specifically:"
    echo "    \` => '"
    echo "    ‘ => '"
    echo "    ´ => '"
    echo "    ’ => '"
    echo '    “ => "'
    echo '    ” => "'
    echo "Conversion takes place by unzipping .xlsx file and manipulating"
    echo "the $stringFile file."
    exit
}

if [ $# -eq 0 ]
then 
    usage
    exit
fi

if [[ $1 == -h* ]]
then
    showHelpAndExit
fi

seenNonXlsx=0

for var in "$@"
do
    if [[ $var == *.xlsx && -f $var ]]
    then
        tmpDir=$var.tmp
        unzip -q -d "$tmpDir" "$var"
        thisStringFile=$tmpDir/$stringFile
        if [ ! -f "$thisStringFile" ]
        then
            echo "ERROR: $stringFile not found in $var. Skipping ..."
            rm -r "$tmpDir"
            continue
        fi
        # Adding the "" after -i is necessary on MacOSX
        # It may not be necessary on Linux machines
        sed -i "" "s/[\`‘´’]/'/g" "$thisStringFile"
        sed -i "" 's/[“”]/"/g' "$thisStringFile"
        cd "$tmpDir"
        zip -rq "../$var" *
        cd ..
        rm -r "$tmpDir"
        echo "Successfully processed $var"
    else
        if [ $seenNonXlsx -eq 0 ]
        then
            echo "WARNING: Program $0 can only work with .xlsx files"
            seenNonXlsx=1
        fi
        echo "Skipping $var ..."
    fi
done