# Script de bash para sacar IOC de Emotet de archivos .doc en formato XML

if [ $# -ne 1 ]; then
	echo "ERROR: Necesita pasarle un archivo como parametro"
fi

TEMP_B64=$(mktemp)
TEMP_OLE=$(mktemp)
TEMP_CODE=$(mktemp)
echo "Extrayendo la información de payload OLE del Word.."
xmlstarlet sel -t -m "//w:docSuppData" -v "w:binData" "$1" > $TEMP_B64
echo

echo "Descomprimiendo el payload..."
python -c 'import sys,zlib,binascii; input = sys.argv[1]; output = sys.argv[2]; f = open(input,"r"); ole_data=zlib.decompress(binascii.a2b_base64(f.read())[0x32:]); f.close(); f = open(output,"a"); f.write(ole_data); f.close();' $TEMP_B64 $TEMP_OLE
echo

echo "Extrayendo el codigo con olevba y sed..."
olevba -c "$TEMP_OLE" | grep -e '"' | sed -re "s/\s*\+\s*//g" -e "s/^[^=]+\s*=\s*//" -e "s/\"//g" | sed -re "s/^.*-e\s*//" | paste -sd "" - > $TEMP_CODE
rm "$TEMP_B64" "$TEMP_OLE"
TEST=$(grep -e "[^A-Za-z0-9=]" "$TEMP_CODE")
echo

if [ -z "$TEST" ]; then
	echo "Al parecer es un codigo en Base64...imprimiendo"
	cat $TEMP_CODE | base64 -d
	echo
else
	echo "Al parecer es el codigo directo..."
	cat $TEMP_CODE
fi

echo 
echo "El codigo esta en el archivo $TEMP_CODE"
