<?xml version="1.0" encoding="UTF-8"?>
<xsl:stylesheet xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
    xmlns="http://schemas.openxmlformats.org/package/2006/relationships"
    xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" version="1.0"
    exclude-result-prefixes="w r">

    <xsl:output indent="no" encoding="UTF-8" version="1.0"/>



    <!-- 
	Le script Python passe dans $tempdir le chemin absolu vers le répertoire temporaire où se trouve le fichier
	XML source.
	La lib Python lxml ne gère pas bien la fonction XPath document() et n'accepte que des chemins
	relatifs à la feuille XSLT, ou des chemins absolus (2ème argument de la fonction document() non géré).
	-->
    <xsl:param name="tempdir"/>
    <xsl:variable name="tempdirSlash" select="translate($tempdir, '\', '/')"/>



    <xsl:variable name="stylesdocument"
        select="document(concat($tempdirSlash,
        '/word/styles.xml'))"/>
    <!--
    <xsl:variable name="stylesdocument" select="document('styles.xml', /)"/>
    -->


    <!-- Copie par défaut. -->
    <xsl:template match="node() | @*">
        <xsl:copy>
            <xsl:apply-templates select="node() | @*"/>
        </xsl:copy>
    </xsl:template>



    <!-- 
    Cette template ajoute des commentaires aux paragraphes mal formatés.
    
    I.   Elle détecte les styles de paragraphes qu'on ne vérifie pas et qui doivent être copiés
         tels quels (auteurs, images...)
        
    II.  Elle affecte ces 4 variables :
         1. Le premier segment de texte du paragraphe, pour vérifier s'il commence par une majuscule,
         2. Le dernier segment de texte du paragraphe, pour vérifier s'il se termine par une
            ponctuation,
         3. Un booléen si le premier segment de texte commence par une majuscule ou non,
         4. Un booléen si le dernier segment de texte se termine par une ponctuation ou non.
         
    III. Les deux dernières variables sont testées simultanément pour insérer un commentaire adapté.
         Un id est donné au commentaire. Il est généré avec la fonction generate-id().
         - Selon les specs XSLT, l'id généré par generate-id() doit commencer par des lettres (et peut
           en contenir ensuite).
         - Selon les specs Office Open XML, l'id des commentaires ne doit pas contenir de lettres.
         ==> On remplace toutes les lettres générées par la fonction generate-id() par un nombre.
         
         Il peut y avoir 3 messages de commentaires :
         1. Paragraphe sans majuscule ni ponctuation finale.
         2. Paragraphe sans ponctuation finale.
         3. Paragraphe sans majuscule.
         Ils sont créés dans le fichier comments.xml lors d'une transformation ultérieure.
         Pour savoir quel message affecter aux commentaires sans correspondance, on se base sur le
         dernier numéro de l'id : 1, 2 ou 3 qui est concaténé à la fin de sa génération.
         - Soit le commentaire existe déjà, il a une correspondance dans comments.xml et dans ce cas
           peu importe son numéro de fin.
         - Soit le commentaire a été créé dans cette transformation, il n'a pas de correspondance
           dans comments.xml, et dans ce cas son id se termine par 1, 2 ou 3 et permet de lui
           affecter un message.
    -->
    <xsl:template match="w:p">
        <xsl:choose>
            <!--
            On ne vérifie pas les paragraphes suivants :
            - Titres
            - Auteurs
            - Bibliographie
            - Tableaux
            - Images
            - Paragraphes de métadonnées (mots-clés, date, pagination...)
            
            On ne connait pas exactement les noms des styles de ces paragraphes dans la mesure
            où ils peuvent être définis par l'utilisateur.
            Il y a un peu plus de certitude lorsque ces styles sont natifs Word (Title...).
            C'est pourquoi le test balaye très large pour éviter de se retrouver avec des 
            commentaires partout.
            -->
            <xsl:when
                test="
                descendant::w:pPr/w:pStyle/@w:val
                = $stylesdocument//w:style[contains(w:name/@w:val, 'auteur')]/@w:styleId
                or descendant::w:pPr/w:pStyle/@w:val
                = $stylesdocument//w:style[contains(w:name/@w:val, 'Auteur')]/@w:styleId
                or descendant::w:pPr/w:pStyle/@w:val
                = $stylesdocument//w:style[contains(w:basedOn/@w:val, 'auteur')]/@w:styleId
                or descendant::w:pPr/w:pStyle/@w:val
                = $stylesdocument//w:style[contains(w:basedOn/@w:val, 'Auteur')]/@w:styleId
                
                or descendant::w:pPr/w:pStyle/@w:val
                = $stylesdocument//w:style[contains(w:name/@w:val, 'bibliographie')]/@w:styleId
                or descendant::w:pPr/w:pStyle/@w:val
                = $stylesdocument//w:style[contains(w:name/@w:val, 'Bibliographie')]/@w:styleId
                or descendant::w:pPr/w:pStyle/@w:val
                = $stylesdocument//w:style[contains(w:basedOn/@w:val, 'bibliographie')]/@w:styleId
                or descendant::w:pPr/w:pStyle/@w:val
                = $stylesdocument//w:style[contains(w:basedOn/@w:val, 'Bibliographie')]/@w:styleId
                
                or descendant::w:pPr/w:pStyle/@w:val
                = $stylesdocument//w:style[contains(w:name/@w:val, 'bibliography')]/@w:styleId
                or descendant::w:pPr/w:pStyle/@w:val
                = $stylesdocument//w:style[contains(w:name/@w:val, 'Bibliography')]/@w:styleId
                or descendant::w:pPr/w:pStyle/@w:val
                = $stylesdocument//w:style[contains(w:basedOn/@w:val, 'bibliography')]/@w:styleId
                or descendant::w:pPr/w:pStyle/@w:val
                = $stylesdocument//w:style[contains(w:basedOn/@w:val, 'Bibliography')]/@w:styleId
                
                or descendant::w:pPr/w:pStyle/@w:val
                = $stylesdocument//w:style[contains(w:name/@w:val, 'code')]/@w:styleId
                or descendant::w:pPr/w:pStyle/@w:val
                = $stylesdocument//w:style[contains(w:name/@w:val, 'Code')]/@w:styleId
                or descendant::w:pPr/w:pStyle/@w:val
                = $stylesdocument//w:style[contains(w:basedOn/@w:val, 'code')]/@w:styleId
                or descendant::w:pPr/w:pStyle/@w:val
                = $stylesdocument//w:style[contains(w:basedOn/@w:val, 'Code')]/@w:styleId
                
                or descendant::w:pPr/w:pStyle/@w:val
                = $stylesdocument//w:style[contains(w:name/@w:val, 'date')]/@w:styleId
                or descendant::w:pPr/w:pStyle/@w:val
                = $stylesdocument//w:style[contains(w:name/@w:val, 'Date')]/@w:styleId
                or descendant::w:pPr/w:pStyle/@w:val
                = $stylesdocument//w:style[contains(w:basedOn/@w:val, 'date')]/@w:styleId
                or descendant::w:pPr/w:pStyle/@w:val
                = $stylesdocument//w:style[contains(w:basedOn/@w:val, 'Date')]/@w:styleId
                
                or descendant::w:pPr/w:pStyle/@w:val
                = $stylesdocument//w:style[contains(w:name/@w:val, 'geographie')]/@w:styleId
                or descendant::w:pPr/w:pStyle/@w:val
                = $stylesdocument//w:style[contains(w:name/@w:val, 'Geographie')]/@w:styleId
                or descendant::w:pPr/w:pStyle/@w:val
                = $stylesdocument//w:style[contains(w:basedOn/@w:val, 'geographie')]/@w:styleId
                or descendant::w:pPr/w:pStyle/@w:val
                = $stylesdocument//w:style[contains(w:basedOn/@w:val, 'Geographie')]/@w:styleId
                
                or descendant::w:pPr/w:pStyle/@w:val
                = $stylesdocument//w:style[contains(w:name/@w:val, 'geography')]/@w:styleId
                or descendant::w:pPr/w:pStyle/@w:val
                = $stylesdocument//w:style[contains(w:name/@w:val, 'Geography')]/@w:styleId
                or descendant::w:pPr/w:pStyle/@w:val
                = $stylesdocument//w:style[contains(w:basedOn/@w:val, 'geography')]/@w:styleId
                or descendant::w:pPr/w:pStyle/@w:val
                = $stylesdocument//w:style[contains(w:basedOn/@w:val, 'Geography')]/@w:styleId
                
                or descendant::w:pPr/w:pStyle/@w:val
                = $stylesdocument//w:style[contains(w:name/@w:val, 'heading')]/@w:styleId
                or descendant::w:pPr/w:pStyle/@w:val
                = $stylesdocument//w:style[contains(w:name/@w:val, 'Heading')]/@w:styleId
                or descendant::w:pPr/w:pStyle/@w:val
                = $stylesdocument//w:style[contains(w:basedOn/@w:val, 'heading')]/@w:styleId
                or descendant::w:pPr/w:pStyle/@w:val
                = $stylesdocument//w:style[contains(w:basedOn/@w:val, 'Heading')]/@w:styleId
                
                or descendant::w:pPr/w:pStyle/@w:val
                = $stylesdocument//w:style[contains(w:name/@w:val, 'langue')]/@w:styleId
                or descendant::w:pPr/w:pStyle/@w:val
                = $stylesdocument//w:style[contains(w:name/@w:val, 'Langue')]/@w:styleId
                or descendant::w:pPr/w:pStyle/@w:val
                = $stylesdocument//w:style[contains(w:basedOn/@w:val, 'langue')]/@w:styleId
                or descendant::w:pPr/w:pStyle/@w:val
                = $stylesdocument//w:style[contains(w:basedOn/@w:val, 'Langue')]/@w:styleId
                
                or descendant::w:pPr/w:pStyle/@w:val
                = $stylesdocument//w:style[contains(w:name/@w:val, 'motscles')]/@w:styleId
                or descendant::w:pPr/w:pStyle/@w:val
                = $stylesdocument//w:style[contains(w:name/@w:val, 'MotsCles')]/@w:styleId
                or descendant::w:pPr/w:pStyle/@w:val
                = $stylesdocument//w:style[contains(w:basedOn/@w:val, 'motscles')]/@w:styleId
                or descendant::w:pPr/w:pStyle/@w:val
                = $stylesdocument//w:style[contains(w:basedOn/@w:val, 'MotsCles')]/@w:styleId
                
                or descendant::w:pPr/w:pStyle/@w:val
                = $stylesdocument//w:style[contains(w:name/@w:val, 'notice')]/@w:styleId
                or descendant::w:pPr/w:pStyle/@w:val
                = $stylesdocument//w:style[contains(w:name/@w:val, 'Notice')]/@w:styleId
                or descendant::w:pPr/w:pStyle/@w:val
                = $stylesdocument//w:style[contains(w:basedOn/@w:val, 'notice')]/@w:styleId
                or descendant::w:pPr/w:pStyle/@w:val
                = $stylesdocument//w:style[contains(w:basedOn/@w:val,'Notice')]/@w:styleId
                
                or descendant::w:pPr/w:pStyle/@w:val
                = $stylesdocument//w:style[contains(w:name/@w:val, 'periode')]/@w:styleId
                or descendant::w:pPr/w:pStyle/@w:val
                = $stylesdocument//w:style[contains(w:name/@w:val, 'Periode')]/@w:styleId
                or descendant::w:pPr/w:pStyle/@w:val
                = $stylesdocument//w:style[contains(w:basedOn/@w:val, 'periode')]/@w:styleId
                or descendant::w:pPr/w:pStyle/@w:val
                = $stylesdocument//w:style[contains(w:basedOn/@w:val, 'Periode')]/@w:styleId
                
                or descendant::w:pPr/w:pStyle/@w:val
                = $stylesdocument//w:style[contains(w:name/@w:val, 'pagination')]/@w:styleId
                or descendant::w:pPr/w:pStyle/@w:val
                = $stylesdocument//w:style[contains(w:name/@w:val, 'Pagination')]/@w:styleId
                or descendant::w:pPr/w:pStyle/@w:val
                = $stylesdocument//w:style[contains(w:basedOn/@w:val, 'pagination')]/@w:styleId
                or descendant::w:pPr/w:pStyle/@w:val
                = $stylesdocument//w:style[contains(w:basedOn/@w:val, 'Pagination')]/@w:styleId
                
                or descendant::w:pPr/w:pStyle/@w:val
                = $stylesdocument//w:style[contains(w:name/@w:val, 'separateur')]/@w:styleId
                or descendant::w:pPr/w:pStyle/@w:val
                = $stylesdocument//w:style[contains(w:name/@w:val, 'Separateur')]/@w:styleId
                or descendant::w:pPr/w:pStyle/@w:val
                = $stylesdocument//w:style[contains(w:basedOn/@w:val, 'separateur')]/@w:styleId
                or descendant::w:pPr/w:pStyle/@w:val
                = $stylesdocument//w:style[contains(w:basedOn/@w:val, 'Separateur')]/@w:styleId
                
                or descendant::w:pPr/w:pStyle/@w:val
                = $stylesdocument//w:style[contains(w:name/@w:val, 'title')]/@w:styleId
                or descendant::w:pPr/w:pStyle/@w:val
                = $stylesdocument//w:style[contains(w:name/@w:val, 'Title')]/@w:styleId
                or descendant::w:pPr/w:pStyle/@w:val
                = $stylesdocument//w:style[contains(w:basedOn/@w:val, 'title')]/@w:styleId
                or descendant::w:pPr/w:pStyle/@w:val
                = $stylesdocument//w:style[contains(w:basedOn/@w:val, 'Title')]/@w:styleId
                
                or descendant::w:pPr/w:pStyle/@w:val
                = $stylesdocument//w:style[contains(w:name/@w:val, 'titre')]/@w:styleId
                or descendant::w:pPr/w:pStyle/@w:val
                = $stylesdocument//w:style[contains(w:name/@w:val, 'Titre')]/@w:styleId
                or descendant::w:pPr/w:pStyle/@w:val
                = $stylesdocument//w:style[contains(w:basedOn/@w:val, 'titre')]/@w:styleId
                or descendant::w:pPr/w:pStyle/@w:val
                = $stylesdocument//w:style[contains(w:basedOn/@w:val, 'Titre')]/@w:styleId
                
                or ancestor::w:tbl
                or not(descendant::w:t)
                ">
                <xsl:copy>
                    <xsl:apply-templates select="node() | @*"/>
                </xsl:copy>
            </xsl:when>

            <xsl:otherwise>
                <xsl:variable name="starttext" select="descendant::w:r[1]//w:t[1]/text()"/>
                <xsl:variable name="endtext" select="descendant::w:r[last()]//w:t[last()]/text()"/>

                <!-- Est-ce que le parahraphe commence par un de ces signes ? -->
                <xsl:variable name="startCap">
                    <xsl:choose>
                        <xsl:when
                            test="starts-with($starttext, 'A')
                            or starts-with($starttext, 'B')
                            or starts-with($starttext, 'C')
                            or starts-with($starttext, 'D')
                            or starts-with($starttext, 'E')
                            or starts-with($starttext, 'F')
                            or starts-with($starttext, 'G')
                            or starts-with($starttext, 'H')
                            or starts-with($starttext, 'I')
                            or starts-with($starttext, 'J')
                            or starts-with($starttext, 'K')
                            or starts-with($starttext, 'L')
                            or starts-with($starttext, 'M')
                            or starts-with($starttext, 'N')
                            or starts-with($starttext, 'O')
                            or starts-with($starttext, 'P')
                            or starts-with($starttext, 'Q')
                            or starts-with($starttext, 'R')
                            or starts-with($starttext, 'S')
                            or starts-with($starttext, 'T')
                            or starts-with($starttext, 'U')
                            or starts-with($starttext, 'V')
                            or starts-with($starttext, 'W')
                            or starts-with($starttext, 'X')
                            or starts-with($starttext, 'Y')
                            or starts-with($starttext, 'Z')
                            or starts-with($starttext, 'À')
                            or starts-with($starttext, 'Á')
                            or starts-with($starttext, 'Â')
                            or starts-with($starttext, 'Ã')
                            or starts-with($starttext, 'Ä')
                            or starts-with($starttext, 'Æ')
                            or starts-with($starttext, 'Ç')
                            or starts-with($starttext, 'È')
                            or starts-with($starttext, 'É')
                            or starts-with($starttext, 'Ê')
                            or starts-with($starttext, 'Ë')
                            or starts-with($starttext, 'Ì')
                            or starts-with($starttext, 'Í')
                            or starts-with($starttext, 'Î')
                            or starts-with($starttext, 'Ï')
                            or starts-with($starttext, 'Ð')
                            or starts-with($starttext, 'Ñ')
                            or starts-with($starttext, 'Ò')
                            or starts-with($starttext, 'Ó')
                            or starts-with($starttext, 'Ø')
                            or starts-with($starttext, 'Ù')
                            or starts-with($starttext, 'Ú')
                            or starts-with($starttext, 'Û')
                            or starts-with($starttext, 'Ü')
                            or starts-with($starttext, 'Ý')
                            or starts-with($starttext, 'Œ')
                            or starts-with($starttext, '1')
                            or starts-with($starttext, '2')
                            or starts-with($starttext, '3')
                            or starts-with($starttext, '4')
                            or starts-with($starttext, '5')
                            or starts-with($starttext, '6')
                            or starts-with($starttext, '7')
                            or starts-with($starttext, '8')
                            or starts-with($starttext, '9')
                            or starts-with($starttext, '0')
                            or starts-with($starttext, '¿')
                            or starts-with($starttext, '¡')
                            or starts-with($starttext, '—')
                            or starts-with($starttext, '-')
                            or starts-with($starttext, '–')
                            or starts-with($starttext, '«')">
                            <xsl:value-of select="'true'"/>
                        </xsl:when>
                        <xsl:otherwise>
                            <xsl:value-of select="'false'"/>
                        </xsl:otherwise>
                    </xsl:choose>
                </xsl:variable>

                <!-- Est-ce que le paragraphe finit par un de ces signes ? -->
                <xsl:variable name="endPunctuation">
                    <xsl:variable name="endsign"
                        select="substring($endtext, string-length($endtext))"/>

                    <xsl:choose>
                        <xsl:when
                            test="$endsign = '.'
                            or $endsign = '?'
                            or $endsign = '!'
                            or $endsign = ':'
                            or $endsign = ';'
                            or $endsign = ','
                            or $endsign = '»'
                            or $endsign = '…'">
                            <xsl:value-of select="'true'"/>
                        </xsl:when>
                        <xsl:otherwise>
                            <xsl:value-of select="'false'"/>
                        </xsl:otherwise>
                    </xsl:choose>
                </xsl:variable>

                <xsl:variable name="commentstyle"
                    select="$stylesdocument//w:style[w:name/@w:val='annotation reference']/@w:styleId"/>

                <xsl:variable name="alphabet"
                    select="'abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ_-'"/>
                <xsl:variable name="numbers" select="'0123456789012345678901234567'"/>


                <!-- Si le paragraphe commence par une majuscule et se termine par une ponctuation -->
                <xsl:choose>
                    <xsl:when test="$startCap = 'true' and $endPunctuation = 'true'">
                        <xsl:copy>
                            <xsl:apply-templates select="node() | @*"/>
                        </xsl:copy>
                    </xsl:when>

                    <xsl:otherwise>
                        <xsl:variable name="id">
                            <xsl:choose>
                                <!-- Ni majuscule, ni ponctuation finale -->
                                <xsl:when test="$startCap = 'false' and $endPunctuation = 'false'">
                                    <xsl:value-of
                                        select="concat(translate(concat(generate-id(), generate-id(..)),
                                        $alphabet, $numbers), '1')"
                                    />
                                </xsl:when>
                                <!-- Une majuscule, mais pas de ponctuation finale -->
                                <xsl:when test="$startCap = 'true' and $endPunctuation = 'false'">
                                    <xsl:value-of
                                        select="concat(translate(concat(generate-id(), generate-id(..)),
                                        $alphabet, $numbers), '2')"
                                    />
                                </xsl:when>
                                <!-- Pas de majuscule, mais ponctuation finale -->
                                <xsl:when test="$startCap = 'false' and $endPunctuation = 'true'">
                                    <xsl:value-of
                                        select="concat(translate(concat(generate-id(), generate-id(..)),
                                        $alphabet, $numbers), '3')"
                                    />
                                </xsl:when>
                            </xsl:choose>
                        </xsl:variable>

                        <xsl:copy>
                            <xsl:apply-templates select="@*"/>
                            <xsl:apply-templates select="descendant::w:pPr"/>

                            <xsl:element name="w:commentRangeStart">
                                <xsl:attribute name="w:id">
                                    <xsl:value-of select="$id"/>
                                </xsl:attribute>
                            </xsl:element>

                            <xsl:apply-templates select="node()[not(self::w:pPr)]"/>

                            <xsl:element name="w:commentRangeEnd">
                                <xsl:attribute name="w:id">
                                    <xsl:value-of select="$id"/>
                                </xsl:attribute>
                            </xsl:element>

                            <xsl:element name="w:r">
                                <xsl:element name="w:rPr">
                                    <xsl:element name="w:rStyle">
                                        <xsl:attribute name="w:val">
                                            <xsl:value-of select="$commentstyle"/>
                                        </xsl:attribute>
                                    </xsl:element>
                                </xsl:element>
                                <xsl:element name="w:commentReference">
                                    <xsl:attribute name="w:id">
                                        <xsl:value-of select="$id"/>
                                    </xsl:attribute>
                                </xsl:element>
                            </xsl:element>
                        </xsl:copy>
                    </xsl:otherwise>
                </xsl:choose>
            </xsl:otherwise>
        </xsl:choose>
    </xsl:template>
</xsl:stylesheet>
