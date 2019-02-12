# Script de bash para sacar IOC de Emotet de archivos .doc en formato XML

hash olevba 2>/dev/null || { echo >&2 "Error: Se necesita instalar el programa 'olevba', favor instalarlo antes de ejecutar el script."; exit 1; }
hash xmlstarlet 2>/dev/null || { echo >&2 "Error: Se necesita instalar el programa 'xmlstarlet', favor instalarlo antes de ejecutar el script."; exit 1; }

if [ $# -ne 1 ]; then
	echo >&2 "ERROR: Necesita pasarle un archivo como parametro"
	exit 1
fi

TEMP_B64=$(mktemp)
TEMP_OLE=$(mktemp)
TEMP_CODE=$(mktemp)
echo "Extrayendo la información de payload OLE del Word.."
xmlstarlet sel -t -m "//w:docSuppData" -v "w:binData" "$1" > $TEMP_B64
echo

sed -i -e "s/\s*//g" "$TEMP_B64"

if [ -z "$(cat $TEMP_B64)" ]; then
	echo >&2 "ERROR: El payload no existe, el archivo esta limpio"
	exit 1
fi

echo "Descomprimiendo el payload..."
python -c 'import sys,zlib,binascii; input = sys.argv[1]; output = sys.argv[2]; f = open(input,"r"); ole_data=zlib.decompress(binascii.a2b_base64(f.read())[0x32:]); f.close(); f = open(output,"a"); f.write(ole_data); f.close();' "$TEMP_B64" "$TEMP_OLE"
echo

echo "Extrayendo el codigo con olevba y sed..."
olevba -c "$TEMP_OLE" | grep -e '"' | sed -re "s/\s*\+\s*//g" -e "s/^[^=]+\s*=\s*//" -e "s/\"//g" | sed -re "s/^.*-e\s*//" | paste -sd "" - | sed -re "s/\s*$//" -e "s/^\s*//" > $TEMP_CODE
rm "$TEMP_B64" "$TEMP_OLE"
TEST=$(cat "$TEMP_CODE" | sed -re "s/[A-Za-z0-9=+\/ ]+//g")
echo

if [ -z "$TEST" ]; then
	echo "Al parecer es un codigo en Base64"
	CODE=$(cat "$TEMP_CODE")
        echo "$CODE" | base64 -d | tr -d '\0' > $TEMP_CODE
	echo
else
	echo "Al parecer es el codigo directo..."
	#CODE=$(cat "$TEMP_CODE")
fi

echo
echo "Realizando una limpieza de las concatenaciones"
echo "Imprimiendo codigo final"
cat -e "$TEMP_CODE" | sed -re "s/'\+'//g" -e "s/([;{}])/\1\n/g"

rm "$TEMP_CODE"

echo 
