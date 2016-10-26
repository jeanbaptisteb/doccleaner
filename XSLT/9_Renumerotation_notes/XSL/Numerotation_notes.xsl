<?xml version="1.0" encoding="UTF-8"?>
<xsl:stylesheet xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
    xmlns:ve="http://schemas.openxmlformats.org/markup-compatibility/2006"
    xmlns:o="urn:schemas-microsoft-com:office:office"
    xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
    xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"
    xmlns:v="urn:schemas-microsoft-com:vml"
    xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
    xmlns:w10="urn:schemas-microsoft-com:office:word"
    xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    xmlns:wne="http://schemas.microsoft.com/office/word/2006/wordml" version="1.0">

    <xsl:output indent="no" encoding="UTF-8" version="1.0"/>



    <xsl:variable name="endOrFootnotes" select="local-name(/node())"/>
    <xsl:variable name="numbers" select="'1234567890_-. '"/>



    <!-- 
	Le script Python passe dans $tempdir le chemin absolu vers le répertoire temporaire où se trouve le fichier
	XML source.
	La lib Python lxml ne gère pas bien la fonction XPath document() et n'accepte que des chemins
	relatifs à la feuille XSLT, ou des chemins absolus (2ème argument de la fonction document() non géré).
	-->
    <xsl:param name="tempdir"/>
    <xsl:variable name="tempdirSlash" select="translate($tempdir, '\', '/')"/>


    <!--
    <xsl:variable name="stylesdocument" select="document(concat($tempdirSlash,'/word/styles.xml'))"/>
    -->
    <xsl:variable name="stylesdocument" select="document('styles.xml', /)"/>




    <!--
    Cette transformation XSLT nettoie les numéros des appels de note au début des notes.
    - Certains numéros de notes sont inscrits en dur, plutôt que d'être incrémentés automatiquement.
    - Certains numéros de notes sont absents.
      ==> La template rétablit les numéros de note automatiques, et supprime les numéros de notes
          manuels le cas échéant.
    
    Une note est encodée en 3 parties comme suit :
    ==============================================================================================
    <w:footnote>
        <w:p> 
    =====  1ère partie : style de la note, immuable  =============================================
            <w:pPr>
                <w:pStyle w:val="Notedebasdepage"/>
            </w:pPr>
    ===== 2ème partie : style du numéro de note, et numéro de note automatique, immuable =========
            <w:r>
                <w:rPr>
                    <w:rStyle w:val="Appelnotedebasdep"/>
                </w:rPr>
                <w:footnoteRef/>
            </w:r>
    ===== 3ème partie : textes de la note =========================================================
            <w:r>
                <w:t>Texte de la note</w:t>
                <w:t>Suite du texte de la note</w:t>
            </w:r>
            <w:r>
                <w:t>Suite du texte de la note.</w:t>
            </w:r> 
        </w:p>
    </w:footnote>
    ==============================================================================================
    
    I. Pour chaque w:p, on va vérifier les 3 parties de la note :
    
    1. On copie w:pPr s'il existe, SINON IL FAUT LE CREER
    
    2. - Si le style du numéro de note est bien formé (s'il existe et qu'il contient une balise de
         numéro de note automatique) : on le copie.
       - S'il n'est pas bien formé : on le construit avec une balise d'appel de note automatique.
         (w:r[w:rPr], avec pour fils  w:r/w:rPr  et w:r/wfootnoteRef)
           
    3. Ensuite, on copie le texte de la note. Cf. template suivante.
       (w:r["qui n'ont pas de descendant '@w:val=Appelnotedebasdep'"])
    
    
    II. Pour chaque w:r :
        Pour les w:r de la troisième partie, ceux qui contiennent le texte de la note, on ne veut
        pas que ce texte commence par un numéro de note manuel.
        Il faut tester toutes les fantaisies d'encodage de Word. Par exemple :
        
        Cas 1 corrigé :
        <w:r>
           <w:t>1. </w:t>
           <w:t>Texte de la première note.</w:t>
        </w:r>
        
        Cas 2 corrigé :
        <w:r>
           <w:t>1</w:t>
           <w:t>2</w:t>
           <w:t> Texte de la douxième note.</w:t>
        </w:r>
        
        Cas 3 corrigé :
        <w:r>
           <w:t>1</w:t>
           <w:t>23</w:t>
           <w:t>4 Texte de la 1234ème note.</w:t>
        </w:r>
        
        Cas 4 corrigé :
        <w:r>
           <w:t>   1</w:t>
        </w:r>
        <w:r>
           <w:t>2</w:t>
           <w:t>2</w:t>
        </w:r>
        <w:r>
           <w:t> La note 122 suit la note</w:t>
        </w:r>
        <w:r>
           <w:t>121 </w:t>
        </w:r>
        <w:r>
           <w:t>et précède la note </w:t>
        </w:r>
        <w:r>
           <w:t>1</w:t>
        </w:r>
        <w:r>
           <w:t>2</w:t>
           <w:t>3</w:t>
        </w:r>
        <w:r>
           <w:t>.</w:t>
        </w:r>
        
        Cas 5 accepté (pas d'espace entre le dernier chiffre et le texte) :
        <w:r>
           <w:t> 3ème place ex-aequo.</w:t>
        </w:r>
        
        Cas 6 accepté :
        <w:r>
           <w:t> L'émission </w:t>
           <w:t>30 </w:t>
           <w:t>millions d'amis.</w:t>
        </w:r>
        
        1. On ne copie pas les w:r dont le texte ne contient que des chiffres, et qui n'est pas
           précédé par un w:r dont le texte ne contient pas que des chiffres.
           ==> Règle complètement les cas 1 et 2, et en partie le cas 3. Pour les situations où
               le numéro de note manuel est 
    
    -->



    <!-- Copie par défaut. -->
    <xsl:template match="node() | @*">
        <xsl:copy>
            <xsl:apply-templates select="node() | @*"/>
        </xsl:copy>
    </xsl:template>



    <xsl:template
        match="w:p[not(descendant::w:separator)][not(descendant::w:continuationSeparator)]">
        <xsl:copy>
            <xsl:copy-of select="@*"/>
            <xsl:apply-templates select="w:pPr"/>

            <xsl:choose>
                <!-- Si le style "appel de note" est bien là et qu'il contient un numéro de note
                    automatique, on le copie -->
                <xsl:when
                    test="(w:r/w:rPr/w:rStyle/@w:val = $stylesdocument//w:style[w:name/@w:val =
                    'footnote reference']/@w:styleId)
                    and w:r[w:rPr/w:rStyle/@w:val = $stylesdocument//w:style[w:name/@w:val =
                    'footnote reference']/@w:styleId]/descendant::w:footnoteRef
                    | w:r/w:rPr/w:rStyle/@w:val = $stylesdocument//w:style[w:name/@w:val =
                    'endnote reference']/@w:styleId
                    and w:r[w:rPr/w:rStyle/@w:val = $stylesdocument//w:style[w:name/@w:val =
                    'endnote reference']/@w:styleId]/descendant::w:endnoteRef
                    ">

                    <xsl:choose>
                        <xsl:when test="$endOrFootnotes = 'footnotes'">
                            <xsl:apply-templates
                                select="w:r[w:rPr/w:rStyle/@w:val = $stylesdocument//w:style[w:name/@w:val =
                                'footnote reference']/@w:styleId]"
                            />
                        </xsl:when>
                        <xsl:otherwise>
                            <xsl:apply-templates
                                select="w:r[w:rPr/w:rStyle/@w:val = $stylesdocument//w:style[w:name/@w:val =
                                'endnote reference']/@w:styleId]"
                            />
                        </xsl:otherwise>
                    </xsl:choose>
                </xsl:when>

                <!-- Sinon, on le crée avec le numéro de note automatique -->
                <xsl:otherwise>
                    <w:r>
                        <xsl:choose>
                            <xsl:when test="$endOrFootnotes = 'footnotes'">
                                <w:rPr>
                                    <w:rStyle
                                        w:val="{$stylesdocument//w:style[w:name/@w:val = 'footnote reference']/@w:styleId}"
                                    />
                                </w:rPr>
                                <w:footnoteRef/>
                            </xsl:when>
                            <xsl:otherwise>
                                <w:rPr>
                                    <w:rStyle
                                        w:val="{$stylesdocument//w:style[w:name/@w:val = 'endnote reference']/@w:styleId}"
                                    />
                                </w:rPr>
                                <w:endnoteRef/>
                            </xsl:otherwise>
                        </xsl:choose>
                    </w:r>
                </xsl:otherwise>
            </xsl:choose>

            <!-- Ensuite, copie de la note -->
            <xsl:choose>
                <xsl:when test="$endOrFootnotes = 'footnotes'">
                    <xsl:apply-templates
                        select="w:r[not(w:rPr/w:rStyle/@w:val = $stylesdocument//w:style[w:name/@w:val =
                        'footnote reference']/@w:styleId)]"
                    />
                </xsl:when>
                <xsl:otherwise>
                    <xsl:apply-templates
                        select="w:r[not(w:rPr/w:rStyle/@w:val = $stylesdocument//w:style[w:name/@w:val =
                        'endnote reference']/@w:styleId)]"
                    />
                </xsl:otherwise>
            </xsl:choose>
        </xsl:copy>
    </xsl:template>


    <!-- CHECKER : si ça ne sélectionne pas les w:r d'appel de notes automatiques (section 2) -->
    <xsl:template
        match="w:r[not(descendant::w:continuationSeparator)][not(descendant::w:separator)]">
        <xsl:variable name="preceding-suppressed-r-node">            
            <xsl:for-each select="
                preceding::w:r
                [not(descendant::w:rStyle/@w:val = $stylesdocument//w:style[w:name/@w:val =
                'footnote reference']/@w:styleId)]
                [ancestor::w:p = current()/ancestor::w:p]
                | preceding::w:r[ancestor::w:p = current()/ancestor::w:p]
                [not(descendant::w:rStyle/@w:val = $stylesdocument//w:style[w:name/@w:val =
                'endnote reference']/@w:styleId)]
                ">
                <xsl:choose>
                    <xsl:when test="translate(descendant::w:t/descendant::text(), $numbers, '') = ''">
                        <xsl:value-of select="'true'"/>
                    </xsl:when>
                    <xsl:otherwise>
                        <xsl:value-of select="'false'"/>
                    </xsl:otherwise>
                </xsl:choose>
            </xsl:for-each>
        </xsl:variable>
        
        <xsl:choose>
            <!-- Si le w:r courant ne contient que des chiffres... -->
            <xsl:when test="translate(descendant::w:t/descendant::text(), $numbers, '') = ''">
                <xsl:choose>
                    <!--
                    ... et qu'il est n'est précédé par rien, ou par des w:r qui ne
                    contiennent que des chiffres : on ne copie pas.
                    -->
                    <xsl:when
                        test="not(contains($preceding-suppressed-r-node, 'false'))
                        or $preceding-suppressed-r-node = ''">
                    </xsl:when>
                    
                    <!-- Sinon on le copie -->
                    <xsl:otherwise>
                        <xsl:copy>
                            <xsl:apply-templates select="node() | @*"/>
                        </xsl:copy>
                    </xsl:otherwise>
                </xsl:choose>
            </xsl:when>
            
            <!-- Sinon : on copie -->
            <xsl:otherwise>
                <xsl:copy>
                    <xsl:apply-templates select="node() | @*"/>
                </xsl:copy>
            </xsl:otherwise>
        </xsl:choose>
    </xsl:template>



    <xsl:template match="w:t">
        <!--
        Ici on cherche s'il n'y a pas un w:t précédent non supprimé
        Mais il faut aussi chercher s'il n'y a pas un w:r précédent non supprimé
        -->
        
        
        <xsl:variable name="preceding-suppressed-r-node">            
            <xsl:for-each select="
                ancestor::w:r/preceding::w:r
                [not(descendant::w:rStyle/@w:val = $stylesdocument//w:style[w:name/@w:val =
                'footnote reference']/@w:styleId)]
                [ancestor::w:p = current()/ancestor::w:p]
                
                | ancestor::w:r/preceding::w:r[ancestor::w:p = current()/ancestor::w:p]
                [not(descendant::w:rStyle/@w:val = $stylesdocument//w:style[w:name/@w:val =
                'endnote reference']/@w:styleId)]
                ">
                <xsl:choose>
                    <xsl:when test="translate(descendant::w:t/descendant::text(), $numbers, '') = ''">
                        <xsl:value-of select="'true'"/>
                    </xsl:when>
                    <xsl:otherwise>
                        <xsl:value-of select="'false'"/>
                    </xsl:otherwise>
                </xsl:choose>
            </xsl:for-each>
        </xsl:variable>
        
        <TEST><xsl:value-of select="$preceding-suppressed-r-node"/></TEST>
        
        <xsl:choose>
            <!--
            Si le w:r parent du w:t courant n'est précédé par rien, ou par des w:r qui ne
            contiennent que des chiffres : on ne copie pas
             
            
            CHECKER : si les w:r précédents le w:r parent ne contiennent pas que des chiffres,
            le w:t courant est copié normalement.
            Sinon, on cherche si ce ne sont pas des numéros de notes manuels.
            -->
            
            
            <xsl:when test="contains($preceding-suppressed-r-node, 'false')
                or $preceding-suppressed-r-node = ''">
                
            </xsl:when>
        </xsl:choose>
        
        
        
        <xsl:variable name="preceding-suppressed-text">
            <xsl:for-each select="preceding::w:t[ancestor::w:r = current()/ancestor::w:r]">
                <xsl:choose>
                    <xsl:when test="translate(descendant::text(), $numbers, '') = ''">
                        <xsl:value-of select="'true'"/>
                    </xsl:when>
                    <xsl:otherwise>
                        <xsl:value-of select="'false'"/>
                    </xsl:otherwise>
                </xsl:choose>
            </xsl:for-each>
        </xsl:variable>
        
        
        

        <xsl:choose>
            <!-- Si le w:t ne contient que des chiffres... -->
            <xsl:when test="translate(descendant::text(), $numbers, '') = ''">
                <xsl:choose>
                    <!--
                    ... et qu'il est n'est précédé par rien, ou par des w:t qui ne
                    contiennent que des chiffres : on ne copie pas.
                    -->
                    <xsl:when
                        test="not(contains($preceding-suppressed-text, 'false'))
                                or $preceding-suppressed-text = ''"/>

                    <!-- Sinon on le copie -->
                    <xsl:otherwise>
                        <xsl:copy>
                            <xsl:apply-templates select="node() | @*"/>
                        </xsl:copy>
                    </xsl:otherwise>
                </xsl:choose>
            </xsl:when>

            <!-- Sinon on le copie -->
            <xsl:otherwise>
                <xsl:copy>
                    <xsl:apply-templates select="node() | @*"/>
                </xsl:copy>
            </xsl:otherwise>
        </xsl:choose>
    </xsl:template>



    <!--
    <xsl:template match="text()[ancestor::w:t]">
        <xsl:variable name="preceding-suppressed-text">
            <xsl:for-each
                select="ancestor::w:r/preceding::w:r[ancestor::w:p =
                    current()/ancestor::w:p]/descendant::w:t">
                <xsl:choose>
                    <xsl:when test="translate(descendant::text(), $numbers, '') = ''">
                        <xsl:value-of select="'true'"/>
                    </xsl:when>
                    <xsl:otherwise>
                        <xsl:value-of select="'false'"/>
                    </xsl:otherwise>
                </xsl:choose>
            </xsl:for-each>
        </xsl:variable>

        <xsl:choose>
            
            Si le texte est précédé de textes qui ont été complètement supprimés 
            (parce que c'était des appels de notes manuels)
            
            <xsl:when
                test="not(contains($preceding-suppressed-text, 'false'))
                and $preceding-suppressed-text != ''">
                <xsl:choose>
                    
                    Si le texte commence par des chiffres, suivis d'un espace :
                    ==> On supprime les chiffres (considérés comme la suite de l'appel de note manuel).
                    
                    <xsl:when
                        test="translate(substring-before(normalize-space(.), ' '), $numbers, '') = ''">
                        <xsl:text> </xsl:text>
                        <xsl:value-of select="substring-after(normalize-space(.), ' ')"/>
                    </xsl:when>
                     Sinon con copie normalement 
                    <xsl:otherwise>
                        <xsl:value-of select="."/>
                    </xsl:otherwise>
                </xsl:choose>
            </xsl:when>
             Sinon on copie normalement 
            <xsl:otherwise>
                <xsl:value-of select="."/>
            </xsl:otherwise>
        </xsl:choose>
    </xsl:template>
    -->
</xsl:stylesheet>
