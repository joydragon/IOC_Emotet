# Script de bash para sacar IOC de Emotet de archivos .doc en formato XML

hash olevba 2>/dev/null || { echo >&2 "Error: Se necesita instalar el programa 'olevba', favor instalarlo antes de ejecutar el script."; exit 1; }
hash xmlstarlet 2>/dev/null || { echo >&2 "Error: Se necesita instalar el programa 'xmlstarlet', favor instalarlo antes de ejecutar el script."; exit 1; }
hash file 2>/dev/null || { echo >&2 "Error: Se necesita instalar el programa 'file', favor instalarlo antes de ejecutar el script."; exit 1; }
hash gzip 2>/dev/null || { echo >&2 "Error: Se necesita instalar el programa 'gzip', favor instalarlo antes de ejecutar el script."; exit 1; }

if [ $# -ne 1 ]; then
	echo >&2 "ERROR: Necesita pasarle un archivo como parametro"
	exit 1
fi

TEMP_CODE=$(mktemp)

function ole_parse1(){
	CODE=$(olevba -c "$1" | grep -e "cmd" | grep -v "binar" | tail -n 1 > "$TEMP_CODE")
	echo >&2 "- Limpiando los \"set\" de CMD"
	SEP="%%%"
	GO=1
	while [ $GO -ne 0 ]; do
		CODE=$(cat "$TEMP_CODE" | grep "set")
		SETX=$(echo "$CODE" | sed -re "s/.*set\s*([^=]+)=([^& ]+).*/\1${SEP}\2/");
		if [ -n "$SETX" ]; then
			P1=$(echo "$SETX" | sed -re "s/(.*)${SEP}.*/\1/")
			P2=$(echo "$SETX" | sed -re "s/.*${SEP}(.*)/\1/")

			echo "$CODE" | sed -r -e "s/set\s*${P1}\s*=\s*${P2}\s*&+//" -e "s/%${P1}%/${P2}/" > "$TEMP_CODE"
		else
			GO=0
		fi
	done
	echo >&2

	echo >&2 "- Buscando y reemplazando codigo de ofuscacion con replace"
	echo >&2
	REPLACE=$(cat "$TEMP_CODE" | sed -re "s/.*replace\('([-0-9]+).*/\1/");
	sed -ire "s/${REPLACE}//g" "$TEMP_CODE";
}

function zlib_parse(){
	ZLIB_AUX=$(cat "$TEMP_CODE" | paste -sd "" - | grep -ie "fromBase64String")
	
	if [ -n "$ZLIB_AUX" ]; then
		echo >&2
		echo >&2 "- Descomprimiendo deflate con base64..."
		echo >&2
		TEMP_ZLIB=$(mktemp)
		echo "$ZLIB_AUX" | sed -re "s/^.*fRoMBAse64stRINg\(+\s*\(*\s*['\"]([^']+)['\"]\).*$/\1/i" | base64 -d > "$TEMP_ZLIB"

		printf "\x1f\x8b\x08\x00\x00\x00\x00\x00\x00\x00" | cat - "$TEMP_ZLIB" | gzip -d 2>/dev/null > "$TEMP_CODE"
		rm "$TEMP_ZLIB"
	fi
}

function format_parse(){
REGEX="\(\"(\{[0-9]+\})+\" *-f[^\(\)]+\)"
FULL_STRING=$(cat "$TEMP_CODE")

while [[ $FULL_STRING =~ $REGEX ]]; do
	
	STRING="${BASH_REMATCH[0]}"
	ORDER=$(echo "$STRING" | sed -re "s/^[^{]+([^'\"]+)['\"]\s*-f\s*(.*)+\).*$/\1/" -e "s/\{//g" )
        FORMAT=$(echo "$STRING" | sed -re "s/^[^{]+([^'\"]+)['\"]\s*-f\s*([^\)]+)+\).*$/\2/" -e "s/'//g")

	IFS='}' read -r -a OR_FORMAT <<< "$ORDER"
        IFS=',' read -r -a AR_FORMAT <<< "$FORMAT"

        i=0
        RES=""
        while [ -n "${AR_FORMAT[$i]}" ]; do
                RES="$RES${AR_FORMAT[${OR_FORMAT[$i]}]}"
                i=$((i+1))
        done

	RES=$(echo $RES | sed -re "s/\//\\\\\//g" -e "s/\{/\\\\\{/g" -e "s/\}/\\\\\}/" -e "s/\(/\\\\\(/" -e "s/\)/\\\\\)/")
	STRING=$(echo $STRING | sed -re "s/\//\\\\\//g" -e "s/\{/\\\\\{/g" -e "s/\}/\\\\\}/" -e "s/\(/\\\\\(/" -e "s/\)/\\\\\)/")
	FULL_STRING=$(echo $FULL_STRING | sed -re "s/$STRING/'$RES'/")
done

echo $FULL_STRING | sed -re "s/\`//g" -e "s/'\+'//g" > "$TEMP_CODE"

SPLIT=$(echo "$FULL_STRING" | sed -re "s/^.*invoke\(['\"](.)['\"].*$/\1/i" -e "s/^.*split['\"]*\(['\"](.)['\"].*$/\1/i")
cat "$TEMP_CODE" |  grep -i "split" | sed -re "s/^.*\(?[\"'](.*)[\"']\)?.\(?[\"']?split.*$/\1/i" -e "s/\\\s*$//" -e "s/${SPLIT}/\n/g" >&2

}

function string_parse(){
	echo >&2 "- Usando metodo San Expedito, para extraer informacion en Base64"
	echo >&2
	# Si hay mas de 1 linea, filtrar las que se repite un mismo caracter por mas de 5 veces seguidas
	DATA=$(strings -n 100 "$1" | grep -v -e "[<>\/ ]" | sed -re "s/(.)\1{4}//g" | grep -ve " xml" | strings -n 100)
	echo -e "$DATA" > $TEMP_CODE
}


function ole_parse3(){
	echo >&2 "- Usando metodo Ave Maria, para extraer informacion en Base64"
	echo >&2
	DATA=$(olevba -c "$1" | grep -ve "VBA FORM" -e "in file" -e "olevba" -e "$1" | grep -e "[A-Za-z0-9+\/]\{20,\}=*" | sed -re "s/^[^A-Za-z0-9]*([A-Za-z0-9+\/]+=*).*$/\1/" | strings -n 100)
	if [ -z "$DATA" ]; then
		echo >&2 "- Datos no encontrados..."
		echo >&2
	else
		echo "$DATA" > "$TEMP_CODE"
	fi
}

function ole_parse2(){
	string_parse "$1"
	if [[ -z "$(cat $TEMP_CODE)" ]] || [[ "$(wc -l $TEMP_CODE)" != "1" ]]; then
		ole_parse3 "$1"
	fi
	if [ -z "$(cat $TEMP_CODE)" ]; then
		echo >&2 "- Extrayendo el codigo con olevba, grep y sed."
		echo >&2
		olevba -c "$1" | grep -e '"' | grep -e "=" | sed -re "s/\s*\+\s*//g" -e "s/^[^=]+\s*=\s*//" -e "s/\"//g" -e "s/GetObject.*$//i" | paste -sd "" - | sed -re "s/^.*-e\s*//" -e "s/^([^=]+=+).*/\1/" -e "s/\s*$//" -e "s/^\s*//" > $TEMP_CODE
	fi
	if [ -z "$(cat $TEMP_CODE)" ]; then
		echo >&2 "- No se pudo sacar informacion."
		echo >&2
		exit 1
	fi

	TEST=$(cat "$TEMP_CODE" | sed -re "s/[A-Za-z0-9=+\/ ]+//g")

	if [ -z "$TEST" ]; then
		echo >&2 "- Al parecer es un codigo en Base64"
		CODE=$(cat "$TEMP_CODE")
		echo "$CODE" | base64 -d | tr -d '\0' > $TEMP_CODE
	else
		echo >&2 "- Al parecer es el codigo directo..."
	fi

	echo >&2
	echo >&2 "- Realizando una limpieza de las concatenaciones e imprimiendo codigo final"
	echo >&2

	CODE=$(cat "$TEMP_CODE")
	echo -e "$CODE" | sed -re "s/'\+'//g" -e "s/\`//g" > "$TEMP_CODE"

	B64_CHK=$(grep "$TEMP_CODE" -ie "Io.COmPReSsiON.DEFlaTestREaM")
	if [ -n "$B64_CHK" ]; then
		echo >&2 "- El codigo esta comprimido dentro del powershell, descomprimiendo..."
		zlib_parse
	fi
}

function xml_parse(){
	TEMP_B64=$(mktemp)
	TEMP_OLE=$(mktemp)
	echo >&2 "- Extrayendo la información de payload OLE del Word.."
	xmlstarlet sel -t -m "//w:docSuppData" -v "w:binData" "$1" > $TEMP_B64
	echo >&2

	sed -i -e "s/\s*//g" "$TEMP_B64"

	if [ -z "$(cat $TEMP_B64)" ]; then
		echo >&2 "ERROR: El payload no existe, el archivo esta limpio"
		exit 1
	fi

	echo >&2 "- Descomprimiendo el payload..."
	python -c 'import sys,zlib,binascii; input = sys.argv[1]; output = sys.argv[2]; f = open(input,"r"); ole_data=zlib.decompress(binascii.a2b_base64(f.read())[0x32:]); f.close(); f = open(output,"a"); f.write(ole_data); f.close();' "$TEMP_B64" "$TEMP_OLE"
	echo >&2

	RES=$(ole_parse2 "$1")

	echo -e "$RES"
}

function doc_parse(){
	echo >&2 "- Extrayendo el codigo con olevba y grep cmd."
	echo >&2
	DOC=$(olevba -c "$1" | grep -e "cmd" | grep -v "binar" | tail -n 1)

	if [ -n "$DOC" ]; then
		RES=$(ole_parse1 "$1")
	else
		RES=$(ole_parse2 "$1")
	fi

	echo -e "$RES"
}

function print_result(){
	echo >&2 "- Imprimiendo el codigo..."
	echo >&2 "------------------------------------------------------------------------------------------------"
	cat "$TEMP_CODE" >&2
	echo >&2
	echo >&2 "------------------------------------------------------------------------------------------------"

	echo >&2
	echo >&2 "- URL de IOC:"
	echo >&2

	FORMAT_CHK=$(grep "$TEMP_CODE" -ie "{[0-9]\+}\s*{[0-9]\+}\s*{[0-9]\+}\s*")
	if [ -n "$FORMAT_CHK" ]; then
		echo >&2 "- Se necesita el filtro extra anti ofuscacion..."
		echo >&2
		format_parse
	else	
		cat "$TEMP_CODE" |  grep -i "split" | sed -re "s/^.*\(?[\"'](.*)[\"']\)?.\(?[\"']?split.*$/\1/i" -e "s/\\\s*$//" -e "s/[,@]/\n/g" >&2
	fi

	rm "$TEMP_CODE"

}

if [ -f "$1" ]; then
	FILETYPE=$(file "$1")
	if [ -n "$(echo "$FILETYPE" | grep -e "XML 1.0 document")" ]; then
		echo >&2 "Revisando el archivo, parece ser un documento Microsoft Office XML"
		echo >&2
		EPOCH=1
		RES=$(xml_parse "$1")

		print_result

	elif [ -n "$(echo "$FILETYPE" | grep -e "Composite Document File V2 Document")" ]; then
		echo >&2  "Revisando el archivo, parece ser un documento .doc con macros"
		echo >&2
		EPOCH=2
		RES=$(doc_parse "$1")

		print_result

	elif [ -n "$(echo "$FILETYPE" | grep -e "Microsoft Word 2007+")" ]; then
		echo >&2 "Revisando el archivo, parece ser un documento .docx con macros"
		echo >&2
		EPOCH=2
		RES=$(doc_parse "$1")

		print_result
	else
		echo >&2 "ERROR: El archivo no tiene el formato requerido (DOC con macros o Microsoft Office XML)"
		exit 1
	fi
else
	echo >&2 "ERROR: el archivo no existe"
	exit 1
fi
