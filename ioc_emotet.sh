# Script de bash para sacar IOC de Emotet de archivos .doc en formato XML

hash olevba 2>/dev/null || { echo >&2 "Error: Se necesita instalar el programa 'olevba', favor instalarlo antes de ejecutar el script."; exit 1; }
hash xmlstarlet 2>/dev/null || { echo >&2 "Error: Se necesita instalar el programa 'xmlstarlet', favor instalarlo antes de ejecutar el script."; exit 1; }
hash file 2>/dev/null || { echo >&2 "Error: Se necesita instalar el programa 'file', favor instalarlo antes de ejecutar el script."; exit 1; }

if [ $# -ne 1 ]; then
	echo >&2 "ERROR: Necesita pasarle un archivo como parametro"
	exit 1
fi

TEMP_CODE=$(mktemp)

function xml_parse(){
	TEMP_B64=$(mktemp)
	TEMP_OLE=$(mktemp)
	echo "- Extrayendo la información de payload OLE del Word.."
	xmlstarlet sel -t -m "//w:docSuppData" -v "w:binData" "$1" > $TEMP_B64
	echo

	sed -i -e "s/\s*//g" "$TEMP_B64"

	if [ -z "$(cat $TEMP_B64)" ]; then
		echo >&2 "ERROR: El payload no existe, el archivo esta limpio"
		exit 1
	fi

	echo "- Descomprimiendo el payload..."
	python -c 'import sys,zlib,binascii; input = sys.argv[1]; output = sys.argv[2]; f = open(input,"r"); ole_data=zlib.decompress(binascii.a2b_base64(f.read())[0x32:]); f.close(); f = open(output,"a"); f.write(ole_data); f.close();' "$TEMP_B64" "$TEMP_OLE"
	echo

	echo "- Extrayendo el codigo con olevba y sed..."
	olevba -c "$TEMP_OLE" | grep -e '"' | sed -re "s/\s*\+\s*//g" -e "s/^[^=]+\s*=\s*//" -e "s/\"//g" | paste -sd "" - | sed -re "s/^.*-e\s*//" | sed -re "s/\s*$//" -e "s/^\s*//" > $TEMP_CODE
	rm "$TEMP_B64" "$TEMP_OLE"
	TEST=$(cat "$TEMP_CODE" | sed -re "s/[A-Za-z0-9=+\/ ]+//g")
	echo

	if [ -z "$TEST" ]; then
		echo "- Al parecer es un codigo en Base64"
		CODE=$(cat "$TEMP_CODE")
		echo "$CODE" | base64 -d | tr -d '\0' > $TEMP_CODE
	else
		echo "- Al parecer es el codigo directo..."
		#CODE=$(cat "$TEMP_CODE")
	fi

	echo
	echo "- Realizando una limpieza de las concatenaciones e imprimiendo codigo final"
	echo

	CODE=$(cat "$TEMP_CODE")
	echo -e "$CODE" | sed -re "s/'\+'//g" -e "s/([;{}])/\1\n/g" > "$TEMP_CODE"
}

function doc_parse(){
	echo "- Extrayendo el codigo con olevba y sed..."
	olevba -c "$1" | grep -e "cmd" | tail -n 1 > "$TEMP_CODE"
	echo

	echo "- Limpiando los \"set\" de CMD"
	SEP="%%%"
	GO=1
	while [ $GO -ne 0 ]; do
	CODE=$(cat "$TEMP_CODE" | grep "set")
	SETX=$(echo "$CODE" | sed -re "s/.*set\s*([^=]+)=([^& ]+).*/\1${SEP}\2/");
	if [ -n "$SETX" ]; then
		P1=$(echo "$SETX" | sed -re "s/(.*)${SEP}.*/\1/")
		P2=$(echo "$SETX" | sed -re "s/.*${SEP}(.*)/\1/")

		echo "$CODE" | sed -r -e "s/set\s*${P1}\s*=\s*${P2}\s*&+//" -e "s/%${P1}%/${P2}/" > $TEMP_CODE
	else
		GO=0
	fi
	done
	echo

	echo "- Buscando y reemplazando codigo de ofuscacion con replace"
	echo
	REPLACE=$(cat "$TEMP_CODE" | sed -re "s/.*replace\('([-0-9]+).*/\1/");
	sed -ire "s/${REPLACE}//g" "$TEMP_CODE";

	echo "- Realizando una limpieza de las concatenaciones e imprimiendo codigo final"
	echo
	sed -i -r -e "s/'\+'//g" -e "s/([;{}])/\1\n/g" "$TEMP_CODE"
}

if [ -f "$1" ]; then
	FILETYPE=$(file "$1")
	if [ -n "$(echo "$FILETYPE" | grep -e "XML 1.0 document")" ]; then
		EPOCH=1
		RES=$(xml_parse "$1")
		echo -e "$RES"
		cat "$TEMP_CODE"
		rm "$TEMP_CODE"
	elif [ -n "$(echo "$FILETYPE" | grep -e "Composite Document File V2 Document")" ]; then
		EPOCH=2
		RES=$(doc_parse "$1")
		echo -e "$RES"
		cat "$TEMP_CODE"
		rm "$TEMP_CODE"
	else
		echo >&2 "ERROR: El archivo no tiene el formato requerido (DOC con macros o Microsoft Office XML)"
		exit 1
	fi
else
	echo >&2 "ERROR: el archivo no existe"
	exit 1
fi
