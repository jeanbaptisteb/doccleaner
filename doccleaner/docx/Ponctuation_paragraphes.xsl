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


    <!-- 
    Index des commentaires déjà existants dans le document, dont on va déterminer l'id max.
    L'id des commentaires est incrémenté d'1 en 1 à partir de 0.
    -->
    <xsl:key name="commentReference" match="//w:commentReference" use="ancestor::*"/>



    <!-- 
    Stocke l'id max des commentaires prééxistants dans le document, déterminé en appelant la
    fonction getCommentReferenceMaxId.
    -->
    <xsl:variable name="maxId">
        <xsl:call-template name="getCommentReferenceMaxId">
            <xsl:with-param name="maxId" select="0"/>
            <xsl:with-param name="index" select="1"/>
            <xsl:with-param name="nbcommentReference" select="count(//w:commentReference)"/>
        </xsl:call-template>
    </xsl:variable>



    <!--
    Détermnine l'id max des commentaires prééxistants dans le document.
    
    En XSLT 1.0 il n'y a pas de boucle, d'array ou de fonction native de comparaison de valeurs.
    Il faut faire une fonction récursive, qui boucle sur l'index des commentaires prééxistants et
    enregistre leur id dans une variable $maxId s'il est plus grand que l'id préalablement
    enregistré.
    -->
    <xsl:template name="getCommentReferenceMaxId">
        <xsl:param name="maxId" select="0"/>
        <xsl:param name="index" select="1"/>
        <xsl:param name="nbcommentReference" select="count(//w:commentReference)"/>

        <xsl:choose>
            <xsl:when test="$index &gt; $nbcommentReference">
                <xsl:value-of select="$maxId"/>
            </xsl:when>

            <xsl:otherwise>
                <xsl:variable name="commentReference" select="key('commentReference', /)[$index]"/>

                <xsl:choose>
                    <xsl:when test="$commentReference/@w:id &gt; $maxId">
                        <xsl:call-template name="getCommentReferenceMaxId">
                            <xsl:with-param name="maxId" select="$commentReference/@w:id"/>
                            <xsl:with-param name="index" select="$index +1"/>
                            <xsl:with-param name="nbcommentReference" select="$nbcommentReference"/>
                        </xsl:call-template>
                    </xsl:when>

                    <xsl:otherwise>
                        <xsl:call-template name="getCommentReferenceMaxId">
                            <xsl:with-param name="maxId" select="$maxId"/>
                            <xsl:with-param name="index" select="$index +1"/>
                            <xsl:with-param name="nbcommentReference" select="$nbcommentReference"/>
                        </xsl:call-template>
                    </xsl:otherwise>
                </xsl:choose>
            </xsl:otherwise>
        </xsl:choose>
    </xsl:template>



    <!-- 
    Index des paragraphes, sur lesquels on va boucler pour déterminer si leur ponctuation est valide
    ou non (présence d'une majuscule et d'un signe de ponctuation finale).
    -->
    <xsl:key name="p" match="//w:p" use="ancestor::*"/>



    <!-- 
    Cette variable stocke un id unique pour chaque paragraphe dont la ponctuation est incorrecte.
    Elle appelle la fonction getPId qui détermine les paragraphes fautifs et renvoit une chaîne de
    caractère avec leurs id.
    Concrètement, $pId peut ressembler à ça : '*d0e842*d0e853*d0e1985'
    
    L'id est généré par la fonction XPath generate-id(noeud). Cette fonction produit un id unique qui
    dépend du noeud passé en argument. Si elle est appelée plusieurs fois sur le même noeud, elle
    produira le même id.
    L'id dépend du parseur. Les specs XSLT précisent qu'il commence forcément par une lettre, ensuite il
    peut contenir une combinaison de lettres et de chiffres de n'importe quelle longueur.
    -->
    <xsl:variable name="pId">
        <xsl:call-template name="getPId">
            <xsl:with-param name="index" select="1"/>
        </xsl:call-template>
    </xsl:variable>



    <!-- 
    Cette fonction boucle sur l'index des paragraphes du document. C'est donc une fonction récursive
    qui appelle un par un les paragraphes de l'index, et concatène si nécessaire une valeur de
    retour dans  la variable $pId.
    
    Pour chaque paragraphe de l'index, la fonction teste sa ponctuation : majuscule au début et
    signe de ponctuation à la fin. Certains styles de paragraphes ne sont pas testés.
    S'il y a un problème, la fonction génère un id pour le paragraphe et le concatène dans $pId.
    -->
    <xsl:template name="getPId">
        <xsl:param name="index" select="1"/>
        <xsl:param name="count" select="count(key('p', /))"/>
        <xsl:param name="pId"/>

        <xsl:choose>
            <xsl:when test="$index &gt; $count">
                <xsl:value-of select="$pId"/>
            </xsl:when>

            <xsl:otherwise>
                <xsl:variable name="currentP" select="key('p', /)[$index]"/>

                <!--
                Le for-each n'itère que sur un élément, ça sert à définir l'élément courant sur
                celui-ci.
                -->
                <xsl:for-each select="$currentP">
                    <xsl:choose>
                        <!--
                        Liste blanche de styles de paragraphes à vérifier :
                        - On sélectionne les paragraphes "normaux" (qui n'ont pas de style)
                        - On sélectionne les résumés, citations, dédicaces, encadrés, erratums, ndla, ndlr,
                          paragraphes sans retrait, questions, remerciements, réponses.
                        - On ne sélectionne pas les paragraphes sans styles qui font partie d'un tableau
                        - On ne sélectionne pas les paragraphes sans styles qui font partie d'une image
                        
                        On ne connait pas exactement les noms des styles de ces paragraphes dans la mesure
                        où ils peuvent être définis par l'utilisateur.
                        Il y a un peu plus de certitude lorsque ces styles sont natifs Word (Title...).
                        C'est pourquoi le test balaye très large pour tenter de les cibler plus
                        précisément.
                        -->
                        <xsl:when
                            test="
                            (not(descendant::w:pStyle)
                            
                            or descendant::w:pPr/w:pStyle/@w:val
                            = $stylesdocument//w:style[contains(w:name/@w:val, 'abstract')]/@w:styleId
                            or descendant::w:pPr/w:pStyle/@w:val
                            = $stylesdocument//w:style[contains(w:name/@w:val, 'Abstract')]/@w:styleId
                            or descendant::w:pPr/w:pStyle/@w:val
                            = $stylesdocument//w:style[contains(w:basedOn/@w:val, 'abstract')]/@w:styleId
                            or descendant::w:pPr/w:pStyle/@w:val
                            = $stylesdocument//w:style[contains(w:basedOn/@w:val, 'Abstract')]/@w:styleId
                            
                            or descendant::w:pPr/w:pStyle/@w:val
                            = $stylesdocument//w:style[contains(w:name/@w:val, 'citation')]/@w:styleId
                            or descendant::w:pPr/w:pStyle/@w:val
                            = $stylesdocument//w:style[contains(w:name/@w:val, 'Citation')]/@w:styleId
                            or descendant::w:pPr/w:pStyle/@w:val
                            = $stylesdocument//w:style[contains(w:basedOn/@w:val, 'citation')]/@w:styleId
                            or descendant::w:pPr/w:pStyle/@w:val
                            = $stylesdocument//w:style[contains(w:basedOn/@w:val, 'Citation')]/@w:styleId
                            
                            or descendant::w:pPr/w:pStyle/@w:val
                            = $stylesdocument//w:style[contains(w:name/@w:val, 'Citationbis')]/@w:styleId
                            or descendant::w:pPr/w:pStyle/@w:val
                            = $stylesdocument//w:style[contains(w:name/@w:val, 'Citation bis')]/@w:styleId
                            or descendant::w:pPr/w:pStyle/@w:val
                            = $stylesdocument//w:style[contains(w:basedOn/@w:val, 'Citationbis')]/@w:styleId
                            or descendant::w:pPr/w:pStyle/@w:val
                            = $stylesdocument//w:style[contains(w:basedOn/@w:val, 'Citation bis')]/@w:styleId
                            
                            or descendant::w:pPr/w:pStyle/@w:val
                            = $stylesdocument//w:style[contains(w:name/@w:val, 'Citationter')]/@w:styleId
                            or descendant::w:pPr/w:pStyle/@w:val
                            = $stylesdocument//w:style[contains(w:name/@w:val, 'Citation ter')]/@w:styleId
                            or descendant::w:pPr/w:pStyle/@w:val
                            = $stylesdocument//w:style[contains(w:basedOn/@w:val, 'Citationter')]/@w:styleId
                            or descendant::w:pPr/w:pStyle/@w:val
                            = $stylesdocument//w:style[contains(w:basedOn/@w:val, 'Citation vter')]/@w:styleId
                            
                            or descendant::w:pPr/w:pStyle/@w:val
                            = $stylesdocument//w:style[contains(w:name/@w:val, 'dedicace')]/@w:styleId
                            or descendant::w:pPr/w:pStyle/@w:val
                            = $stylesdocument//w:style[contains(w:name/@w:val, 'Dedicace')]/@w:styleId
                            or descendant::w:pPr/w:pStyle/@w:val
                            = $stylesdocument//w:style[contains(w:basedOn/@w:val, 'dedicace')]/@w:styleId
                            or descendant::w:pPr/w:pStyle/@w:val
                            = $stylesdocument//w:style[contains(w:basedOn/@w:val, 'Dedicace')]/@w:styleId
                            
                            or descendant::w:pPr/w:pStyle/@w:val
                            = $stylesdocument//w:style[contains(w:name/@w:val, 'encadre')]/@w:styleId
                            or descendant::w:pPr/w:pStyle/@w:val
                            = $stylesdocument//w:style[contains(w:name/@w:val, 'Encadre')]/@w:styleId
                            or descendant::w:pPr/w:pStyle/@w:val
                            = $stylesdocument//w:style[contains(w:basedOn/@w:val, 'encadre')]/@w:styleId
                            or descendant::w:pPr/w:pStyle/@w:val
                            = $stylesdocument//w:style[contains(w:basedOn/@w:val, 'Encadre')]/@w:styleId
                            
                            or descendant::w:pPr/w:pStyle/@w:val
                            = $stylesdocument//w:style[contains(w:name/@w:val, 'erratum')]/@w:styleId
                            or descendant::w:pPr/w:pStyle/@w:val
                            = $stylesdocument//w:style[contains(w:name/@w:val, 'Erratum')]/@w:styleId
                            or descendant::w:pPr/w:pStyle/@w:val
                            = $stylesdocument//w:style[contains(w:basedOn/@w:val, 'erratum')]/@w:styleId
                            or descendant::w:pPr/w:pStyle/@w:val
                            = $stylesdocument//w:style[contains(w:basedOn/@w:val, 'Erratum')]/@w:styleId
                            
                            or descendant::w:pPr/w:pStyle/@w:val
                            = $stylesdocument//w:style[contains(w:name/@w:val, 'NDLR')]/@w:styleId
                            or descendant::w:pPr/w:pStyle/@w:val
                            = $stylesdocument//w:style[contains(w:name/@w:val, 'NDLA')]/@w:styleId
                            or descendant::w:pPr/w:pStyle/@w:val
                            = $stylesdocument//w:style[contains(w:basedOn/@w:val, 'NDLR')]/@w:styleId
                            or descendant::w:pPr/w:pStyle/@w:val
                            = $stylesdocument//w:style[contains(w:basedOn/@w:val, 'NDLA')]/@w:styleId
                            
                            or descendant::w:pPr/w:pStyle/@w:val
                            = $stylesdocument//w:style[contains(w:name/@w:val, 'Paragraphesansretrait')]/@w:styleId
                            or descendant::w:pPr/w:pStyle/@w:val
                            = $stylesdocument//w:style[contains(w:name/@w:val, 'Paragraphe sans retrait')]/@w:styleId
                            or descendant::w:pPr/w:pStyle/@w:val
                            = $stylesdocument//w:style[contains(w:basedOn/@w:val, 'Paragraphesansretrait')]/@w:styleId
                            or descendant::w:pPr/w:pStyle/@w:val
                            = $stylesdocument//w:style[contains(w:basedOn/@w:val, 'Paragraphe sans retrait')]/@w:styleId
                            
                            or descendant::w:pPr/w:pStyle/@w:val
                            = $stylesdocument//w:style[contains(w:name/@w:val, 'question')]/@w:styleId
                            or descendant::w:pPr/w:pStyle/@w:val
                            = $stylesdocument//w:style[contains(w:name/@w:val, 'Question')]/@w:styleId
                            or descendant::w:pPr/w:pStyle/@w:val
                            = $stylesdocument//w:style[contains(w:basedOn/@w:val, 'question')]/@w:styleId
                            or descendant::w:pPr/w:pStyle/@w:val
                            = $stylesdocument//w:style[contains(w:basedOn/@w:val, 'Question')]/@w:styleId
                            
                            or descendant::w:pPr/w:pStyle/@w:val
                            = $stylesdocument//w:style[contains(w:name/@w:val, 'quote')]/@w:styleId
                            or descendant::w:pPr/w:pStyle/@w:val
                            = $stylesdocument//w:style[contains(w:name/@w:val, 'Quote')]/@w:styleId
                            or descendant::w:pPr/w:pStyle/@w:val
                            = $stylesdocument//w:style[contains(w:basedOn/@w:val, 'quote')]/@w:styleId
                            or descendant::w:pPr/w:pStyle/@w:val
                            = $stylesdocument//w:style[contains(w:basedOn/@w:val, 'Quote')]/@w:styleId
                            
                            or descendant::w:pPr/w:pStyle/@w:val
                            = $stylesdocument//w:style[contains(w:name/@w:val, 'remerciements')]/@w:styleId
                            or descendant::w:pPr/w:pStyle/@w:val
                            = $stylesdocument//w:style[contains(w:name/@w:val, 'Remerciements')]/@w:styleId
                            or descendant::w:pPr/w:pStyle/@w:val
                            = $stylesdocument//w:style[contains(w:basedOn/@w:val, 'remerciements')]/@w:styleId
                            or descendant::w:pPr/w:pStyle/@w:val
                            = $stylesdocument//w:style[contains(w:basedOn/@w:val, 'Remerciements')]/@w:styleId
                            
                            or descendant::w:pPr/w:pStyle/@w:val
                            = $stylesdocument//w:style[contains(w:name/@w:val, 'reponse')]/@w:styleId
                            or descendant::w:pPr/w:pStyle/@w:val
                            = $stylesdocument//w:style[contains(w:name/@w:val, 'Reponse')]/@w:styleId
                            or descendant::w:pPr/w:pStyle/@w:val
                            = $stylesdocument//w:style[contains(w:basedOn/@w:val, 'reponse')]/@w:styleId
                            or descendant::w:pPr/w:pStyle/@w:val
                            = $stylesdocument//w:style[contains(w:basedOn/@w:val, 'Reponse')]/@w:styleId
                            
                            or descendant::w:pPr/w:pStyle/@w:val
                            = $stylesdocument//w:style[contains(w:name/@w:val, 'resume')]/@w:styleId
                            or descendant::w:pPr/w:pStyle/@w:val
                            = $stylesdocument//w:style[contains(w:name/@w:val, 'Resume')]/@w:styleId
                            or descendant::w:pPr/w:pStyle/@w:val
                            = $stylesdocument//w:style[contains(w:basedOn/@w:val, 'resume')]/@w:styleId
                            or descendant::w:pPr/w:pStyle/@w:val
                            = $stylesdocument//w:style[contains(w:basedOn/@w:val, 'Resume')]/@w:styleId
                            )
                            
                            and not(descendant::w:pPr/descendant::w:numPr)
                            and not(ancestor::w:tbl)
                            and descendant::w:t
                            ">

                            <xsl:variable name="starttext"
                                select="descendant::w:r[1]//w:t[1]/text()"/>
                            <xsl:variable name="endtext"
                                select="descendant::w:r[last()]//w:t[last()]/text()"/>

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

                            <!--
                            Si le paragraphe commence par une majuscule et se termine par une
                            ponctuation
                            -->
                            <xsl:choose>
                                <xsl:when test="$startCap = 'true' and $endPunctuation = 'true'">
                                    <xsl:call-template name="getPId">
                                        <xsl:with-param name="index" select="$index +1"/>
                                        <xsl:with-param name="count" select="$count"/>
                                        <xsl:with-param name="pId" select="$pId"/>
                                    </xsl:call-template>
                                </xsl:when>

                                <xsl:otherwise>
                                    <xsl:call-template name="getPId">
                                        <xsl:with-param name="index" select="$index +1"/>
                                        <xsl:with-param name="count" select="$count"/>
                                        <xsl:with-param name="pId">
                                            <xsl:value-of select="concat($pId, '*', generate-id(.))"
                                            />
                                        </xsl:with-param>
                                    </xsl:call-template>
                                </xsl:otherwise>
                            </xsl:choose>
                        </xsl:when>
                        
                        <xsl:otherwise>
                            <xsl:call-template name="getPId">
                                <xsl:with-param name="index" select="$index +1"/>
                                <xsl:with-param name="count" select="$count"/>
                                <xsl:with-param name="pId" select="$pId"/>
                            </xsl:call-template>
                        </xsl:otherwise>
                    </xsl:choose>
                </xsl:for-each>
            </xsl:otherwise>
        </xsl:choose>
    </xsl:template>



    <!-- Copie par défaut. -->
    <xsl:template match="node() | @*">
        <xsl:copy>
            <xsl:apply-templates select="node() | @*"/>
        </xsl:copy>
    </xsl:template>



    <xsl:template match="w:p">
        <xsl:variable name="currentId" select="generate-id(.)"/>

        <!-- 
        Quand l'id du paragraphe courant est contenu dans $pId, c'est qu'il y a un problème de
        ponctuation (majuscule ou ponctuation finale manquante).
        On copie donc le paragraphe en insérant des balises de commentaire.
        -->
        <xsl:choose>
            <xsl:when test="contains($pId, $currentId)">

                <xsl:variable name="id">
                    <!-- 
                    Franchement, c'est un peu compliqué à expliquer.
                    
                    Alors, $pId est une chaîne de caractères qui peut ressembler à ça :
                    '*d0e842*d0e853*d0e1985'
                    
                    Chaque id correspond à un paragraphe dont la ponctuation pose problème, et
                    auquel il faut rajouter un commentaire.
                    
                    Pour déterminer l'id qu'on va donner à ce commentaire (numérotés de 0 à x) :
                    - on regarde sa position dans $pId (en calculant la différence entre la longueur
                    de la chaîne avant l'id du paragraphe, moins la longueur de la chaîne avant l'id
                    du paragraphe à laquelle on a enlevé les '*'. Ca donne 0 pour le premier
                    paragraphe.).
                    - On ajoute $maxId qui est l'id max des paragraphes qui possèdent déjà un
                    commentaire préalablement à cette transformation.
                    -->
                    <xsl:value-of
                        select="string-length(substring-before($pId, $currentId)) -
                        (string-length(translate(substring-before($pId, $currentId), '*', ''))) + $maxId"
                    />
                </xsl:variable>

                <xsl:variable name="commentstyle"
                    select="$stylesdocument//w:style[w:name/@w:val='annotation reference']/@w:styleId"/>

                <xsl:variable name="alphabet"
                    select="'abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ_-'"/>

                <!--
                Les specs docx demandent qu'un des id soit un nombre (hexadécimal) à 8 chiffres...
                
                L'id généré par generate-id dépend du parseur. Les specs XSLT précisent qu'il
                commence forcément par une lettre, ensuite il peut contenir une combinaison de
                lettres et de chiffres de n'importe quelle longueur.
                On génère donc 8 id sur des noeuds différents, qu'on concatène, dont on supprime les
                lettres, et dont on ne prend que les 8 premiers caractères du résultat.
                -->
                <xsl:variable name="hexId">
                    <xsl:value-of
                        select="
                        substring(
                            translate(
                                concat(
                                    generate-id(.),
                                    generate-id(*),
                                    generate-id(preceding::*[1]),
                                    generate-id(following::*[1]),
                                    generate-id(preceding::*[2]),
                                    generate-id(following::*[2]),
                                    generate-id(..),
                                    generate-id(../..)
                                ),
                            $alphabet, ''),
                        1, 8)"
                    />
                </xsl:variable>

                <!-- 
                On copie le paragraphe en ajoutant des balises de commentaire.
                -->
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
                        <xsl:attribute name="w:rsidR">
                            <xsl:value-of select="$hexId"/>
                        </xsl:attribute>
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
            </xsl:when>

            <!-- 
            Si l'id du paragraphe n'est pas présent dans $pId, il n'y a aucun problème de
            ponctuation, on le copie à l'identique.
            -->
            <xsl:otherwise>
                <xsl:copy>
                    <xsl:apply-templates select="node() | @*"/>
                </xsl:copy>
            </xsl:otherwise>
        </xsl:choose>
    </xsl:template>
</xsl:stylesheet>
