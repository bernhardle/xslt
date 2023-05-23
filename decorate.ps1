#
#	(c) ___, ___ (2023)
#
#	Version:
#		2023-05-03:	Erstellt.
#	Original:
#		
#	Verweise:
#		-/-
#	Verwendet:
#		-/-
#
param([String] $DataDir = $pwd, [String] $stylesheet = 'Decorate.xslt', [Boolean] $verbose = $false, [Boolean] $debug = $false) 
#
Clear-Host
#
Add-Type -Path "C:\WINDOWS\Microsoft.Net\assembly\GAC_MSIL\System.Xml\v4.0_4.0.0.0__b77a5c561934e089\System.Xml.dll"
Add-Type -Path "C:\WINDOWS\Microsoft.Net\assembly\GAC_MSIL\System.IO\v4.0_4.0.0.0__b03f5f7f11d50a3a\System.IO.dll"
Add-Type -Path "C:\WINDOWS\Microsoft.Net\assembly\GAC_MSIL\System.Drawing\v4.0_4.0.0.0__b03f5f7f11d50a3a\System.Drawing.dll"
Add-Type -Path "C:\WINDOWS\Microsoft.Net\assembly\GAC_MSIL\System.Windows.Forms\v4.0_4.0.0.0__b77a5c561934e089\System.Windows.Forms.dll"
#
#	To allow transformations including the 'document()' function.
#
[AppContext]::SetSwitch('Switch.System.Xml.AllowDefaultResolver', $true)
#
[Xml] $script:transform = @"
<?xml version="1.0" encoding="UTF-8"?>
<xsl:stylesheet 
	version="1.0"
	xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
	xmlns:a="http://purl.oclc.org/ooxml/drawingml/main" 
	xmlns:p="http://purl.oclc.org/ooxml/presentationml/main" 
	xmlns:r="http://purl.oclc.org/ooxml/officeDocument/relationships"
	xmlns:q="http://schemas.openxmlformats.org/package/2006/relationships"
	xmlns:w="http://schemas.openxmlformats.org/presentationml/2006/main"
	exclude-result-prefixes="a p r q w">
	<!--

	-->
	<xsl:param name="Decorate.FirstImageIndex" select="1" />
	<xsl:param name="Decorate.FirstRelationIndex" select="2" />
	<xsl:param name="Decorate.TagDataLoadFile" />
	
	<xsl:variable name="Decorate.TagData" select="document(`$Decorate.TagDataLoadFile)/Objects/Object" />
	<!--

	-->
	<xsl:output encoding="UTF-8" method="xml" standalone="yes" indent="yes" />
	<!--
		CustomData und Extensions werden nicht mitgenommen
	-->
	<xsl:template match="a:extLst|p:custDataLst" />

	<xsl:template name="extraRels" xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
		<xsl:param name="startId" />
		<xsl:for-each select="`$Decorate.TagData">
			<Relationship Id="{concat ('rId', (`$startId + position () - 1))}" Type="http://purl.oclc.org/ooxml/officeDocument/relationships/image" Target="{concat ('../media/image', (`$Decorate.FirstImageIndex + position() - 1), '.jpeg')}"/>
		</xsl:for-each>
	</xsl:template>

	<xsl:template name="recRels">
		<xsl:param name="rels" />
		<xsl:param name="maxId" select="0" />
		<xsl:choose>
			<xsl:when test="`$rels">
				<xsl:apply-templates select="`$rels [1]" />
				<xsl:call-template name="recRels">
					<xsl:with-param name="rels" select="`$rels [position() > 1]" />
					<xsl:with-param name="maxId">
						<xsl:variable name="rId" select="number (substring-after (`$rels [1]/@Id, 'rId'))" />
						<xsl:choose>
							<xsl:when test="`$rId > `$maxId">
								<xsl:value-of select="`$rId" />
							</xsl:when>
							<xsl:otherwise>
								<xsl:value-of select="`$maxId" />
							</xsl:otherwise>
						</xsl:choose>
					</xsl:with-param>
				</xsl:call-template>
			</xsl:when>
			<xsl:otherwise>
				<xsl:comment>
					<xsl:text>
				
			*** Ende der mitgenommenen rels, max. rId=</xsl:text><xsl:value-of select="`$maxId" /><xsl:text> ***
			
	</xsl:text>
				</xsl:comment>
				<xsl:call-template name="extraRels">
					<xsl:with-param name="startId" select="`$maxId + 1" />
				</xsl:call-template>
				<xsl:comment>
					<xsl:text>
				
			*** Ende der neuen rels. ***
			
	</xsl:text>
				</xsl:comment>
			</xsl:otherwise>
		</xsl:choose>
	</xsl:template>

	<xsl:template match="q:Relationships">
		<xsl:copy>
			<xsl:call-template name="recRels">
				<xsl:with-param name="rels" select="q:Relationship [not(@Type='http://purl.oclc.org/ooxml/officeDocument/relationships/tags')]" />
			</xsl:call-template>
		</xsl:copy>
	</xsl:template>
	<!--
			
	-->
	<xsl:template name="kachel">
		<xsl:param name="xpos" select="0" />
		<xsl:param name="ypos" select="0" />
		<xsl:param name="line1" select="'Text1'" />
		<xsl:param name="line2" select="'Text2'" />
		<xsl:param name="line3" select="'Text3'" />
		<xsl:param name="line4" select="'Text4'" />
		<xsl:param name="pieces" select="0" />
		<xsl:param name="new" select="0" />
		<xsl:param name="check" select="0" />
		<xsl:param name="imgrel" />
		<xsl:variable name="yincr" select="592835" />
		<xsl:variable name="xincr" select="1457652" />
		<xsl:variable name="boxlinewidth" select="6350" />
		<!-- Weißer Hintergrund -->
		<p:sp>
			<p:nvSpPr>
				<p:cNvPr id="83" name="{concat ('Hintergrund an Position x=', `$xpos, ' y=', `$ypos)}"/>
				<p:cNvSpPr/>
				<p:nvPr/>
			</p:nvSpPr>
			<p:spPr>
				<a:xfrm>
					<a:off x="{279473 + (`$xpos - 1) * `$xincr}" y="{1049454 + (`$ypos -1) * `$yincr}"/>
					<a:ext cx="1363980" cy="502920"/>
				</a:xfrm>
				<a:prstGeom prst="rect">
					<a:avLst/>
				</a:prstGeom>
				<a:solidFill>
					<a:srgbClr val="FFFFFF"/>
				</a:solidFill>
				<a:ln w="{`$boxlinewidth}">
					<a:noFill/>
				</a:ln>
			</p:spPr>
			<p:style>
				<a:lnRef idx="2">
					<a:schemeClr val="accent1">
						<a:shade val="50%"/>
					</a:schemeClr>
				</a:lnRef>
				<a:fillRef idx="1">
					<a:schemeClr val="accent1"/>
				</a:fillRef>
				<a:effectRef idx="0">
					<a:schemeClr val="accent1"/>
				</a:effectRef>
				<a:fontRef idx="minor">
					<a:schemeClr val="lt1"/>
				</a:fontRef>
			</p:style>
			<p:txBody>
				<a:bodyPr lIns="72000" tIns="72000" rIns="72000" bIns="72000" rtlCol="0" anchor="ctr"/>
				<a:lstStyle/>
				<a:p>
					<a:pPr algn="ctr"/>
					<a:endParaRPr lang="en-US" sz="1600" dirty="0" err="1">
						<a:solidFill>
							<a:srgbClr val="000000"/>
						</a:solidFill>
					</a:endParaRPr>
				</a:p>
			</p:txBody>
		</p:sp>
		<!-- Symbolfoto -->
		<p:pic>
				<p:nvPicPr>
					<p:cNvPr id="12" name="{concat('Symbolfoto an Position x=', `$xpos, ' y=', `$ypos)}"/>
					<p:cNvPicPr>
						<a:picLocks noChangeAspect="1"/>
					</p:cNvPicPr>
					<p:nvPr/>
				</p:nvPicPr>
				<p:blipFill>
					<a:blip r:embed="{concat ('rId', `$imgrel)}"/>
					<a:stretch>
						<a:fillRect/>
					</a:stretch>
				</p:blipFill>
				<p:spPr>
					<a:xfrm>
						<a:off x="{350000 + (`$xpos - 1) * `$xincr}" y="{1100000 + (`$ypos - 1) * `$yincr}"/>
						<a:ext cx="600000" cy="400000"/>
					</a:xfrm>
					<a:prstGeom prst="rect">
						<a:avLst/>
					</a:prstGeom>
				</p:spPr>
			</p:pic>
		<!-- Textfeld mit vier Zeilen rechts -->
		<p:sp>
			<p:nvSpPr>
				<p:cNvPr id="87" name="{concat('Textfeldan an Position x=', `$xpos, ' y=', `$ypos)}"/>
				<p:cNvSpPr/>
				<p:nvPr/>
			</p:nvSpPr>
			<p:spPr>
				<a:xfrm>
					<a:off x="{1040000 + (`$xpos - 1) * `$xincr}" y="{1088000 + (`$ypos -1) * `$yincr}"/>
					<a:ext cx="594000" cy="449580"/>
				</a:xfrm>
				<a:prstGeom prst="rect">
					<a:avLst/>
				</a:prstGeom>
				<a:solidFill>
					<a:srgbClr val="FFFFFF"/>
				</a:solidFill>
				<a:ln w="6350" cmpd="sng">
					<a:noFill/>
					<a:prstDash val="solid"/>
				</a:ln>
			</p:spPr>
			<p:style>
				<a:lnRef idx="2">
					<a:schemeClr val="accent1">
						<a:shade val="50%"/>
					</a:schemeClr>
				</a:lnRef>
				<a:fillRef idx="1">
					<a:schemeClr val="accent1"/>
				</a:fillRef>
				<a:effectRef idx="0">
					<a:schemeClr val="accent1"/>
				</a:effectRef>
				<a:fontRef idx="minor">
					<a:schemeClr val="lt1"/>
				</a:fontRef>
			</p:style>
			<p:txBody>
				<a:bodyPr lIns="72000" tIns="72000" rIns="72000" bIns="72000" rtlCol="0" anchor="ctr"/>
				<a:lstStyle/>
				<a:p>
					<a:pPr marL="0" marR="0" lvl="0" indent="0" algn="ctr" defTabSz="914400" rtl="0" eaLnBrk="1" fontAlgn="auto" latinLnBrk="0" hangingPunct="1">
						<a:lnSpc>
							<a:spcPct val="100%"/>
						</a:lnSpc>
						<a:spcBef>
							<a:spcPts val="0"/>
						</a:spcBef>
						<a:spcAft>
							<a:spcPts val="0"/>
						</a:spcAft>
						<a:buClrTx/>
						<a:buSzTx/>
						<a:buFontTx/>
						<a:buNone/>
						<a:tabLst/>
						<a:defRPr/>
					</a:pPr>
					<a:r>
						<a:rPr lang="en-US" sz="650" b="1" dirty="0">
							<a:ln>
								<a:noFill/>
							</a:ln>
							<a:solidFill>
								<a:srgbClr val="000000"/>
							</a:solidFill>
							<a:effectLst/>
							<a:uLnTx/>
							<a:uFillTx/>
							<a:latin typeface="Arial"/>
						</a:rPr>
						<a:t><xsl:value-of select="`$line1" /></a:t>
					</a:r>
				</a:p>
				<a:p>
					<a:pPr marL="0" marR="0" lvl="0" indent="0" algn="ctr" defTabSz="914400" rtl="0" eaLnBrk="1" fontAlgn="auto" latinLnBrk="0" hangingPunct="1">
						<a:lnSpc>
							<a:spcPct val="100%"/>
						</a:lnSpc>
						<a:spcBef>
							<a:spcPts val="0"/>
						</a:spcBef>
						<a:spcAft>
							<a:spcPts val="0"/>
						</a:spcAft>
						<a:buClrTx/>
						<a:buSzTx/>
						<a:buFontTx/>
						<a:buNone/>
						<a:tabLst/>
						<a:defRPr/>
					</a:pPr>
					<a:r>
						<a:rPr lang="en-US" sz="650" b="1" dirty="0">
							<a:solidFill>
								<a:srgbClr val="000000"/>
							</a:solidFill>
							<a:latin typeface="Arial"/>
						</a:rPr>
						<a:t><xsl:value-of select="`$line2" /></a:t>
					</a:r>
				</a:p>
				<a:p>
					<a:pPr marL="0" marR="0" lvl="0" indent="0" algn="ctr" defTabSz="914400" rtl="0" eaLnBrk="1" fontAlgn="auto" latinLnBrk="0" hangingPunct="1">
						<a:lnSpc>
							<a:spcPct val="100%"/>
						</a:lnSpc>
						<a:spcBef>
							<a:spcPts val="0"/>
						</a:spcBef>
						<a:spcAft>
							<a:spcPts val="0"/>
						</a:spcAft>
						<a:buClrTx/>
						<a:buSzTx/>
						<a:buFontTx/>
						<a:buNone/>
						<a:tabLst/>
						<a:defRPr/>
					</a:pPr>
					<a:r>
						<a:rPr lang="en-US" sz="650" b="1" dirty="0">
							<a:ln>
								<a:noFill/>
							</a:ln>
							<a:solidFill>
								<a:srgbClr val="000000"/>
							</a:solidFill>
							<a:effectLst/>
							<a:uLnTx/>
							<a:uFillTx/>
							<a:latin typeface="Arial"/>
						</a:rPr>
						<a:t><xsl:value-of select="`$line3" /></a:t>
					</a:r>
				</a:p>
				<a:p>
					<a:pPr marL="0" marR="0" lvl="0" indent="0" algn="ctr" defTabSz="914400" rtl="0" eaLnBrk="1" fontAlgn="auto" latinLnBrk="0" hangingPunct="1">
						<a:lnSpc>
							<a:spcPct val="100%"/>
						</a:lnSpc>
						<a:spcBef>
							<a:spcPts val="0"/>
						</a:spcBef>
						<a:spcAft>
							<a:spcPts val="0"/>
						</a:spcAft>
						<a:buClrTx/>
						<a:buSzTx/>
						<a:buFontTx/>
						<a:buNone/>
						<a:tabLst/>
						<a:defRPr/>
					</a:pPr>
					<a:r>
						<a:rPr lang="en-US" sz="650" b="1" dirty="0">
							<a:ln>
								<a:noFill/>
							</a:ln>
							<a:solidFill>
								<a:srgbClr val="000000"/>
							</a:solidFill>
							<a:effectLst/>
							<a:uLnTx/>
							<a:uFillTx/>
							<a:latin typeface="Arial"/>
						</a:rPr>
						<a:t><xsl:value-of select="`$line4" /></a:t>
					</a:r>
				</a:p>
			</p:txBody>
		</p:sp>
		<!-- Roter Rahmen -->
		<p:sp>
			<p:nvSpPr>
				<p:cNvPr id="33" name="{concat ('Roter Rahmen an Position x=', `$xpos, ' y=', `$ypos)}"/>
				<p:cNvSpPr/>
				<p:nvPr/>
			</p:nvSpPr>
			<p:spPr>
				<a:xfrm>
					<a:off x="{279473 + (`$xpos - 1) * `$xincr}" y="{1049454 + (`$ypos -1) * `$yincr}"/>
					<a:ext cx="1363980" cy="502920"/>
				</a:xfrm>
				<a:prstGeom prst="rect">
					<a:avLst/>
				</a:prstGeom>
				<a:noFill/>
				<a:ln w="{`$boxlinewidth}">
					<a:solidFill>
						<a:srgbClr val="FF0000"/>
					</a:solidFill>
				</a:ln>
			</p:spPr>
			<p:style>
				<a:lnRef idx="2">
					<a:schemeClr val="accent1">
						<a:shade val="50%"/>
					</a:schemeClr>
				</a:lnRef>
				<a:fillRef idx="1">
					<a:schemeClr val="accent1"/>
				</a:fillRef>
				<a:effectRef idx="0">
					<a:schemeClr val="accent1"/>
				</a:effectRef>
				<a:fontRef idx="minor">
					<a:schemeClr val="lt1"/>
				</a:fontRef>
			</p:style>
			<p:txBody>
				<a:bodyPr lIns="72000" tIns="72000" rIns="72000" bIns="72000" rtlCol="0" anchor="ctr"/>
				<a:lstStyle/>
				<a:p>
					<a:pPr algn="ctr"/>
					<a:endParaRPr lang="en-US" sz="1600" dirty="0" err="1">
						<a:solidFill>
							<a:srgbClr val="000000"/>
						</a:solidFill>
					</a:endParaRPr>
				</a:p>
			</p:txBody>
		</p:sp>
		<!-- Gruener Haken -->
		<xsl:if test="`$check > 0">
			<p:sp>
			<p:nvSpPr>
				<p:cNvPr id="277" name="{concat('Gruener Haken an Position x=', `$xpos, ' y=', `$ypos)}"/>
				<p:cNvSpPr>
					<a:spLocks noChangeAspect="1"/>
				</p:cNvSpPr>
				<p:nvPr/>
			</p:nvSpPr>
			<p:spPr>
				<a:xfrm>
					<a:off x="{1604666 + (`$xpos - 1) * `$xincr}" y="{996615 + (`$ypos -1) * `$yincr}"/>
					<a:ext cx="105923" cy="108000"/>
				</a:xfrm>
				<a:custGeom>
					<a:avLst/>
					<a:gdLst>
						<a:gd name="connsiteX0" fmla="*/ 385513 w 436430"/>
						<a:gd name="connsiteY0" fmla="*/ 0 h 334024"/>
						<a:gd name="connsiteX1" fmla="*/ 436430 w 436430"/>
						<a:gd name="connsiteY1" fmla="*/ 50917 h 334024"/>
						<a:gd name="connsiteX2" fmla="*/ 153965 w 436430"/>
						<a:gd name="connsiteY2" fmla="*/ 333382 h 334024"/>
						<a:gd name="connsiteX3" fmla="*/ 153680 w 436430"/>
						<a:gd name="connsiteY3" fmla="*/ 333097 h 334024"/>
						<a:gd name="connsiteX4" fmla="*/ 152753 w 436430"/>
						<a:gd name="connsiteY4" fmla="*/ 334024 h 334024"/>
						<a:gd name="connsiteX5" fmla="*/ 0 w 436430"/>
						<a:gd name="connsiteY5" fmla="*/ 181272 h 334024"/>
						<a:gd name="connsiteX6" fmla="*/ 50917 w 436430"/>
						<a:gd name="connsiteY6" fmla="*/ 130354 h 334024"/>
						<a:gd name="connsiteX7" fmla="*/ 153038 w 436430"/>
						<a:gd name="connsiteY7" fmla="*/ 232475 h 334024"/>
					</a:gdLst>
					<a:ahLst/>
					<a:cxnLst>
						<a:cxn ang="0">
							<a:pos x="connsiteX0" y="connsiteY0"/>
						</a:cxn>
						<a:cxn ang="0">
							<a:pos x="connsiteX1" y="connsiteY1"/>
						</a:cxn>
						<a:cxn ang="0">
							<a:pos x="connsiteX2" y="connsiteY2"/>
						</a:cxn>
						<a:cxn ang="0">
							<a:pos x="connsiteX3" y="connsiteY3"/>
						</a:cxn>
						<a:cxn ang="0">
							<a:pos x="connsiteX4" y="connsiteY4"/>
						</a:cxn>
						<a:cxn ang="0">
							<a:pos x="connsiteX5" y="connsiteY5"/>
						</a:cxn>
						<a:cxn ang="0">
							<a:pos x="connsiteX6" y="connsiteY6"/>
						</a:cxn>
						<a:cxn ang="0">
							<a:pos x="connsiteX7" y="connsiteY7"/>
						</a:cxn>
					</a:cxnLst>
					<a:rect l="l" t="t" r="r" b="b"/>
					<a:pathLst>
						<a:path w="436430" h="334024">
							<a:moveTo>
								<a:pt x="385513" y="0"/>
							</a:moveTo>
							<a:lnTo>
								<a:pt x="436430" y="50917"/>
							</a:lnTo>
							<a:lnTo>
								<a:pt x="153965" y="333382"/>
							</a:lnTo>
							<a:lnTo>
								<a:pt x="153680" y="333097"/>
							</a:lnTo>
							<a:lnTo>
								<a:pt x="152753" y="334024"/>
							</a:lnTo>
							<a:lnTo>
								<a:pt x="0" y="181272"/>
							</a:lnTo>
							<a:lnTo>
								<a:pt x="50917" y="130354"/>
							</a:lnTo>
							<a:lnTo>
								<a:pt x="153038" y="232475"/>
							</a:lnTo>
							<a:close/>
						</a:path>
					</a:pathLst>
				</a:custGeom>
				<a:solidFill>
					<a:srgbClr val="92D050"/>
				</a:solidFill>
				<a:ln w="9525" cap="flat" cmpd="sng" algn="ctr">
					<a:noFill/>
					<a:prstDash val="solid"/>
				</a:ln>
				<a:effectLst/>
			</p:spPr>
			<p:style>
				<a:lnRef idx="2">
					<a:schemeClr val="accent1">
						<a:shade val="50%"/>
					</a:schemeClr>
				</a:lnRef>
				<a:fillRef idx="1">
					<a:schemeClr val="accent1"/>
				</a:fillRef>
				<a:effectRef idx="0">
					<a:schemeClr val="accent1"/>
				</a:effectRef>
				<a:fontRef idx="minor">
					<a:schemeClr val="lt1"/>
				</a:fontRef>
			</p:style>
			<p:txBody>
				<a:bodyPr rtlCol="0" anchor="ctr"/>
				<a:lstStyle/>
				<a:p>
					<a:pPr marL="0" marR="0" lvl="0" indent="0" algn="ctr" defTabSz="914400" rtl="0" eaLnBrk="1" fontAlgn="auto" latinLnBrk="0" hangingPunct="1">
						<a:lnSpc>
							<a:spcPct val="100%"/>
						</a:lnSpc>
						<a:spcBef>
							<a:spcPts val="0"/>
						</a:spcBef>
						<a:spcAft>
							<a:spcPts val="0"/>
						</a:spcAft>
						<a:buClrTx/>
						<a:buSzTx/>
						<a:buFontTx/>
						<a:buNone/>
						<a:tabLst/>
						<a:defRPr/>
					</a:pPr>
					<a:endParaRPr kumimoji="0" lang="en-US" sz="1292" b="0" i="0" u="none" strike="noStrike" kern="1200" cap="none" spc="0" normalizeH="0" baseline="0%" noProof="0" err="1">
						<a:ln>
							<a:noFill/>
						</a:ln>
						<a:solidFill>
							<a:srgbClr val="FFFFFF"/>
						</a:solidFill>
						<a:effectLst/>
						<a:uLnTx/>
						<a:uFillTx/>
						<a:latin typeface="Arial" panose="020B0604020202020204" pitchFamily="34" charset="0"/>
						<a:ea typeface="+mn-ea"/>
						<a:cs typeface="Arial" panose="020B0604020202020204" pitchFamily="34" charset="0"/>
					</a:endParaRPr>
				</a:p>
			</p:txBody>
		</p:sp>
		</xsl:if>
		<!-- Gelber Kasten -->
		<p:sp>
			<p:nvSpPr>
				<p:cNvPr id="145" name="{concat('Gelbes Feld an Position x=', `$xpos, ' y=', `$ypos)}"/>
				<p:cNvSpPr/>
				<p:nvPr/>
			</p:nvSpPr>
			<p:spPr>
				<a:xfrm>
					<a:off x="{944318 + (`$xpos - 1) * `$xincr}" y="{966022 + (`$ypos - 1) * `$yincr}"/>
					<a:ext cx="612000" cy="154800"/>
				</a:xfrm>
				<a:prstGeom prst="rect">
					<a:avLst/>
				</a:prstGeom>
				<a:solidFill>
					<a:srgbClr val="FFFF00"/>
				</a:solidFill>
				<a:ln w="6350" cmpd="sng">
					<a:solidFill>
						<a:srgbClr val="EBEBEB"/>
					</a:solidFill>
					<a:prstDash val="solid"/>
				</a:ln>
			</p:spPr>
			<p:style>
				<a:lnRef idx="2">
					<a:schemeClr val="accent1">
						<a:shade val="50%"/>
					</a:schemeClr>
				</a:lnRef>
				<a:fillRef idx="1">
					<a:schemeClr val="accent1"/>
				</a:fillRef>
				<a:effectRef idx="0">
					<a:schemeClr val="accent1"/>
				</a:effectRef>
				<a:fontRef idx="minor">
					<a:schemeClr val="lt1"/>
				</a:fontRef>
			</p:style>
			<p:txBody>
				<a:bodyPr lIns="72000" tIns="72000" rIns="72000" bIns="72000" rtlCol="0" anchor="ctr"/>
				<a:lstStyle/>
				<a:p>
					<a:pPr marL="0" marR="0" lvl="0" indent="0" algn="ctr" defTabSz="914400" rtl="0" eaLnBrk="1" fontAlgn="auto" latinLnBrk="0" hangingPunct="1">
						<a:lnSpc>
							<a:spcPct val="100%"/>
						</a:lnSpc>
						<a:spcBef>
							<a:spcPts val="0"/>
						</a:spcBef>
						<a:spcAft>
							<a:spcPts val="0"/>
						</a:spcAft>
						<a:buClrTx/>
						<a:buSzTx/>
						<a:buFontTx/>
						<a:buNone/>
						<a:tabLst/>
						<a:defRPr/>
					</a:pPr>
					<a:r>
						<a:rPr lang="en-US" sz="650" b="1" dirty="0">
							<a:ln>
								<a:noFill/>
							</a:ln>
							<a:solidFill>
								<a:srgbClr val="000000"/>
							</a:solidFill>
							<a:effectLst/>
							<a:uLnTx/>
							<a:uFillTx/>
							<a:latin typeface="Arial"/>
						</a:rPr>
						<a:t><xsl:value-of select="`$pieces" /><xsl:text> pcs.</xsl:text></a:t>
					</a:r>
				</a:p>
			</p:txBody>
		</p:sp>
		<!-- Roter Kasten "NEW" -->
		<xsl:if test="`$new > 0">
			<p:sp>
				<p:nvSpPr>
					<p:cNvPr id="145" name="{concat('Rotes Feld an Position x=', `$xpos, ' y=', `$ypos)}"/>
					<p:cNvSpPr/>
					<p:nvPr/>
				</p:nvSpPr>
				<p:spPr>
					<a:xfrm>
						<a:off x="{944318 + (`$xpos - 1) * `$xincr}" y="{966022 + (`$ypos - 1) * `$yincr}"/>
						<a:ext cx="612000" cy="154800"/>
					</a:xfrm>
					<a:prstGeom prst="rect">
						<a:avLst/>
					</a:prstGeom>
					<a:solidFill>
						<a:srgbClr val="FF0000"/>
					</a:solidFill>
					<a:ln w="6350" cmpd="sng">
						<a:solidFill>
							<a:srgbClr val="EBEBEB"/>
						</a:solidFill>
						<a:prstDash val="solid"/>
					</a:ln>
				</p:spPr>
				<p:style>
					<a:lnRef idx="2">
						<a:schemeClr val="accent1">
							<a:shade val="50%"/>
						</a:schemeClr>
					</a:lnRef>
					<a:fillRef idx="1">
						<a:schemeClr val="accent1"/>
					</a:fillRef>
					<a:effectRef idx="0">
						<a:schemeClr val="accent1"/>
					</a:effectRef>
					<a:fontRef idx="minor">
						<a:schemeClr val="lt1"/>
					</a:fontRef>
				</p:style>
				<p:txBody>
					<a:bodyPr lIns="72000" tIns="72000" rIns="72000" bIns="72000" rtlCol="0" anchor="ctr"/>
					<a:lstStyle/>
					<a:p>
						<a:pPr marL="0" marR="0" lvl="0" indent="0" algn="ctr" defTabSz="914400" rtl="0" eaLnBrk="1" fontAlgn="auto" latinLnBrk="0" hangingPunct="1">
							<a:lnSpc>
								<a:spcPct val="100%"/>
							</a:lnSpc>
							<a:spcBef>
								<a:spcPts val="0"/>
							</a:spcBef>
							<a:spcAft>
								<a:spcPts val="0"/>
							</a:spcAft>
							<a:buClrTx/>
							<a:buSzTx/>
							<a:buFontTx/>
							<a:buNone/>
							<a:tabLst/>
							<a:defRPr/>
						</a:pPr>
						<a:r>
							<a:rPr lang="en-US" sz="650" b="1" dirty="0">
								<a:ln>
									<a:noFill/>
								</a:ln>
								<a:solidFill>
									<a:srgbClr val="FFFFFF"/>
								</a:solidFill>
								<a:effectLst/>
								<a:uLnTx/>
								<a:uFillTx/>
								<a:latin typeface="Arial"/>
							</a:rPr>
							<a:t>New</a:t>
						</a:r>
					</a:p>
				</p:txBody>
			</p:sp>
		</xsl:if>
	</xsl:template>

	<xsl:template match="w:spTree">
		<xsl:message>
			<xsl:text>

	++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
	
	Diese Datei ist nicht im Strict Open XML Format angelegt und wird daher nicht verarbeitet.

	++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
	
</xsl:text>
		</xsl:message>
		<xsl:copy>
			<xsl:apply-templates select="@* | node()" />
		</xsl:copy>
	</xsl:template>

	<xsl:template match="p:spTree">
		<xsl:copy>
			<xsl:comment>
				<xsl:text>
				
			*** Begin original portion of shape tree ***
			
</xsl:text>
			</xsl:comment>
			<xsl:apply-templates select="@* | node()" />
			<xsl:comment>
				<xsl:text>
				
			*** End original portion of shape tree ***
			
</xsl:text>
			</xsl:comment>
			<xsl:comment>
				<xsl:text>
				
			*** Begin added portion of shape tree ***
			
</xsl:text>
			</xsl:comment>
			<xsl:for-each select="`$Decorate.TagData">
				<xsl:call-template name="kachel">
					<xsl:with-param name="xpos" select="translate (substring(Property [@Name='position'], 1, 1),'ABCDEFGH','12345678')" />
					<xsl:with-param name="ypos" select="substring(Property [@Name='position'], 2, 1)" />
					<xsl:with-param name="line1" select="Property [@Name='line1']" />
					<xsl:with-param name="line2" select="Property [@Name='line2']" />
					<xsl:with-param name="line3" select="Property [@Name='line3']" />
					<xsl:with-param name="line4" select="Property [@Name='line4']" />
					<xsl:with-param name="pieces" select="number(Property [@Name='pieces'])" />
					<xsl:with-param name="check">
						<xsl:choose>
							<xsl:when test="Property [@Name='check'] = 'WAHR'">
								<xsl:value-of select="1" />
							</xsl:when>
							<xsl:otherwise>
								<xsl:value-of select="0" />
							</xsl:otherwise>
						</xsl:choose>
					</xsl:with-param>
					<xsl:with-param name="new">
						<xsl:choose>
							<xsl:when test="Property [@Name='new'] = 'WAHR'">
								<xsl:value-of select="1" />
							</xsl:when>
							<xsl:otherwise>
								<xsl:value-of select="0" />
							</xsl:otherwise>
						</xsl:choose>
					</xsl:with-param>
					<xsl:with-param name="imgrel" select="`$Decorate.FirstRelationIndex + position() - 1" />
				</xsl:call-template>
			</xsl:for-each>
			<!-- Legende gruppiert -->
			<p:grpSp>
				<p:nvGrpSpPr>
					<p:cNvPr id="785" name="Gruppieren 784"/>
					<p:cNvGrpSpPr/>
					<p:nvPr/>
				</p:nvGrpSpPr>
				<p:grpSpPr>
					<a:xfrm>
						<a:off x="4896359" y="6351201"/>
						<a:ext cx="3641798" cy="428539"/>
						<a:chOff x="5189149" y="6076963"/>
						<a:chExt cx="3641798" cy="428539"/>
					</a:xfrm>
				</p:grpSpPr>
				<p:sp>
					<p:nvSpPr>
						<p:cNvPr id="249" name="Textfeld 248"/>
						<p:cNvSpPr txBox="1"/>
						<p:nvPr/>
					</p:nvSpPr>
					<p:spPr>
						<a:xfrm>
							<a:off x="7212214" y="6220624"/>
							<a:ext cx="1618733" cy="283906"/>
						</a:xfrm>
						<a:prstGeom prst="rect">
							<a:avLst/>
						</a:prstGeom>
						<a:noFill/>
					</p:spPr>
					<p:txBody>
						<a:bodyPr wrap="square" lIns="72000" tIns="72000" rIns="72000" bIns="72000" rtlCol="0">
							<a:spAutoFit/>
						</a:bodyPr>
						<a:lstStyle/>
						<a:p>
							<a:pPr marL="0" marR="0" lvl="0" indent="0" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" fontAlgn="auto" latinLnBrk="0" hangingPunct="1">
								<a:lnSpc>
									<a:spcPct val="100%"/>
								</a:lnSpc>
								<a:spcBef>
									<a:spcPts val="0"/>
								</a:spcBef>
								<a:spcAft>
									<a:spcPts val="0"/>
								</a:spcAft>
								<a:buClrTx/>
								<a:buSzTx/>
								<a:buFontTx/>
								<a:buNone/>
								<a:tabLst/>
								<a:defRPr/>
							</a:pPr>
							<a:r>
								<a:rPr kumimoji="0" lang="en-US" sz="900" b="0" i="0" u="none" strike="noStrike" kern="1200" cap="none" spc="0" normalizeH="0" baseline="0%" noProof="0" dirty="0">
									<a:ln>
										<a:noFill/>
									</a:ln>
									<a:solidFill>
										<a:prstClr val="black"/>
									</a:solidFill>
									<a:effectLst/>
									<a:uLnTx/>
									<a:uFillTx/>
									<a:latin typeface="Arial"/>
									<a:ea typeface="+mn-ea"/>
									<a:cs typeface="+mn-cs"/>
								</a:rPr>
								<a:t>Introduced </a:t>
							</a:r>
							<a:r>
								<a:rPr lang="en-US" sz="900" dirty="0">
									<a:solidFill>
										<a:prstClr val="black"/>
									</a:solidFill>
									<a:latin typeface="Arial"/>
								</a:rPr>
								<a:t>in previous month</a:t>
							</a:r>
							<a:endParaRPr kumimoji="0" lang="en-US" sz="900" b="0" i="0" u="none" strike="noStrike" kern="1200" cap="none" spc="0" normalizeH="0" baseline="0%" noProof="0" dirty="0">
								<a:ln>
									<a:noFill/>
								</a:ln>
								<a:solidFill>
									<a:prstClr val="black"/>
								</a:solidFill>
								<a:effectLst/>
								<a:uLnTx/>
								<a:uFillTx/>
								<a:latin typeface="Arial"/>
								<a:ea typeface="+mn-ea"/>
								<a:cs typeface="+mn-cs"/>
							</a:endParaRPr>
						</a:p>
					</p:txBody>
				</p:sp>
				<p:sp>
					<p:nvSpPr>
						<p:cNvPr id="106" name="symbol_Check"/>
						<p:cNvSpPr>
							<a:spLocks noChangeAspect="1"/>
						</p:cNvSpPr>
						<p:nvPr/>
					</p:nvSpPr>
					<p:spPr>
						<a:xfrm>
							<a:off x="5189149" y="6305239"/>
							<a:ext cx="105923" cy="108000"/>
						</a:xfrm>
						<a:custGeom>
							<a:avLst/>
							<a:gdLst>
								<a:gd name="connsiteX0" fmla="*/ 385513 w 436430"/>
								<a:gd name="connsiteY0" fmla="*/ 0 h 334024"/>
								<a:gd name="connsiteX1" fmla="*/ 436430 w 436430"/>
								<a:gd name="connsiteY1" fmla="*/ 50917 h 334024"/>
								<a:gd name="connsiteX2" fmla="*/ 153965 w 436430"/>
								<a:gd name="connsiteY2" fmla="*/ 333382 h 334024"/>
								<a:gd name="connsiteX3" fmla="*/ 153680 w 436430"/>
								<a:gd name="connsiteY3" fmla="*/ 333097 h 334024"/>
								<a:gd name="connsiteX4" fmla="*/ 152753 w 436430"/>
								<a:gd name="connsiteY4" fmla="*/ 334024 h 334024"/>
								<a:gd name="connsiteX5" fmla="*/ 0 w 436430"/>
								<a:gd name="connsiteY5" fmla="*/ 181272 h 334024"/>
								<a:gd name="connsiteX6" fmla="*/ 50917 w 436430"/>
								<a:gd name="connsiteY6" fmla="*/ 130354 h 334024"/>
								<a:gd name="connsiteX7" fmla="*/ 153038 w 436430"/>
								<a:gd name="connsiteY7" fmla="*/ 232475 h 334024"/>
							</a:gdLst>
							<a:ahLst/>
							<a:cxnLst>
								<a:cxn ang="0">
									<a:pos x="connsiteX0" y="connsiteY0"/>
								</a:cxn>
								<a:cxn ang="0">
									<a:pos x="connsiteX1" y="connsiteY1"/>
								</a:cxn>
								<a:cxn ang="0">
									<a:pos x="connsiteX2" y="connsiteY2"/>
								</a:cxn>
								<a:cxn ang="0">
									<a:pos x="connsiteX3" y="connsiteY3"/>
								</a:cxn>
								<a:cxn ang="0">
									<a:pos x="connsiteX4" y="connsiteY4"/>
								</a:cxn>
								<a:cxn ang="0">
									<a:pos x="connsiteX5" y="connsiteY5"/>
								</a:cxn>
								<a:cxn ang="0">
									<a:pos x="connsiteX6" y="connsiteY6"/>
								</a:cxn>
								<a:cxn ang="0">
									<a:pos x="connsiteX7" y="connsiteY7"/>
								</a:cxn>
							</a:cxnLst>
							<a:rect l="l" t="t" r="r" b="b"/>
							<a:pathLst>
								<a:path w="436430" h="334024">
									<a:moveTo>
										<a:pt x="385513" y="0"/>
									</a:moveTo>
									<a:lnTo>
										<a:pt x="436430" y="50917"/>
									</a:lnTo>
									<a:lnTo>
										<a:pt x="153965" y="333382"/>
									</a:lnTo>
									<a:lnTo>
										<a:pt x="153680" y="333097"/>
									</a:lnTo>
									<a:lnTo>
										<a:pt x="152753" y="334024"/>
									</a:lnTo>
									<a:lnTo>
										<a:pt x="0" y="181272"/>
									</a:lnTo>
									<a:lnTo>
										<a:pt x="50917" y="130354"/>
									</a:lnTo>
									<a:lnTo>
										<a:pt x="153038" y="232475"/>
									</a:lnTo>
									<a:close/>
								</a:path>
							</a:pathLst>
						</a:custGeom>
						<a:solidFill>
							<a:srgbClr val="92D050"/>
						</a:solidFill>
						<a:ln w="9525" cap="flat" cmpd="sng" algn="ctr">
							<a:noFill/>
							<a:prstDash val="solid"/>
						</a:ln>
						<a:effectLst/>
					</p:spPr>
					<p:style>
						<a:lnRef idx="2">
							<a:schemeClr val="accent1">
								<a:shade val="50%"/>
							</a:schemeClr>
						</a:lnRef>
						<a:fillRef idx="1">
							<a:schemeClr val="accent1"/>
						</a:fillRef>
						<a:effectRef idx="0">
							<a:schemeClr val="accent1"/>
						</a:effectRef>
						<a:fontRef idx="minor">
							<a:schemeClr val="lt1"/>
						</a:fontRef>
					</p:style>
					<p:txBody>
						<a:bodyPr rtlCol="0" anchor="ctr"/>
						<a:lstStyle/>
						<a:p>
							<a:pPr marL="0" marR="0" lvl="0" indent="0" algn="ctr" defTabSz="914400" rtl="0" eaLnBrk="1" fontAlgn="auto" latinLnBrk="0" hangingPunct="1">
								<a:lnSpc>
									<a:spcPct val="100%"/>
								</a:lnSpc>
								<a:spcBef>
									<a:spcPts val="0"/>
								</a:spcBef>
								<a:spcAft>
									<a:spcPts val="0"/>
								</a:spcAft>
								<a:buClrTx/>
								<a:buSzTx/>
								<a:buFontTx/>
								<a:buNone/>
								<a:tabLst/>
								<a:defRPr/>
							</a:pPr>
							<a:endParaRPr kumimoji="0" lang="en-US" sz="1292" b="0" i="0" u="none" strike="noStrike" kern="1200" cap="none" spc="0" normalizeH="0" baseline="0%" noProof="0" err="1">
								<a:ln>
									<a:noFill/>
								</a:ln>
								<a:solidFill>
									<a:srgbClr val="FFFFFF"/>
								</a:solidFill>
								<a:effectLst/>
								<a:uLnTx/>
								<a:uFillTx/>
								<a:latin typeface="Arial" panose="020B0604020202020204" pitchFamily="34" charset="0"/>
								<a:ea typeface="+mn-ea"/>
								<a:cs typeface="Arial" panose="020B0604020202020204" pitchFamily="34" charset="0"/>
							</a:endParaRPr>
						</a:p>
					</p:txBody>
				</p:sp>
				<p:sp>
					<p:nvSpPr>
						<p:cNvPr id="109" name="Textfeld 108"/>
						<p:cNvSpPr txBox="1"/>
						<p:nvPr/>
					</p:nvSpPr>
					<p:spPr>
						<a:xfrm>
							<a:off x="5274243" y="6221596"/>
							<a:ext cx="1515767" cy="283906"/>
						</a:xfrm>
						<a:prstGeom prst="rect">
							<a:avLst/>
						</a:prstGeom>
						<a:noFill/>
					</p:spPr>
					<p:txBody>
						<a:bodyPr wrap="square" lIns="72000" tIns="72000" rIns="72000" bIns="72000" rtlCol="0">
							<a:spAutoFit/>
						</a:bodyPr>
						<a:lstStyle/>
						<a:p>
							<a:pPr marL="0" marR="0" lvl="0" indent="0" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" fontAlgn="auto" latinLnBrk="0" hangingPunct="1">
								<a:lnSpc>
									<a:spcPct val="100%"/>
								</a:lnSpc>
								<a:spcBef>
									<a:spcPts val="0"/>
								</a:spcBef>
								<a:spcAft>
									<a:spcPts val="0"/>
								</a:spcAft>
								<a:buClrTx/>
								<a:buSzTx/>
								<a:buFontTx/>
								<a:buNone/>
								<a:tabLst/>
								<a:defRPr/>
							</a:pPr>
							<a:r>
								<a:rPr kumimoji="0" lang="en-US" sz="900" b="0" i="0" u="none" strike="noStrike" kern="1200" cap="none" spc="0" normalizeH="0" baseline="0%" noProof="0" dirty="0">
									<a:ln>
										<a:noFill/>
									</a:ln>
									<a:solidFill>
										<a:prstClr val="black"/>
									</a:solidFill>
									<a:effectLst/>
									<a:uLnTx/>
									<a:uFillTx/>
									<a:latin typeface="Arial"/>
									<a:ea typeface="+mn-ea"/>
									<a:cs typeface="+mn-cs"/>
								</a:rPr>
								<a:t>Implemented</a:t>
							</a:r>
						</a:p>
					</p:txBody>
				</p:sp>
				<p:sp>
					<p:nvSpPr>
						<p:cNvPr id="30" name="Rechteck 29"/>
						<p:cNvSpPr/>
						<p:nvPr/>
					</p:nvSpPr>
					<p:spPr>
						<a:xfrm>
							<a:off x="6104823" y="6298884"/>
							<a:ext cx="219075" cy="114300"/>
						</a:xfrm>
						<a:prstGeom prst="rect">
							<a:avLst/>
						</a:prstGeom>
						<a:solidFill>
							<a:srgbClr val="FFFF00"/>
						</a:solidFill>
						<a:ln w="6350" cmpd="sng">
							<a:solidFill>
								<a:srgbClr val="EBEBEB"/>
							</a:solidFill>
							<a:prstDash val="solid"/>
						</a:ln>
					</p:spPr>
					<p:style>
						<a:lnRef idx="2">
							<a:schemeClr val="accent1">
								<a:shade val="50%"/>
							</a:schemeClr>
						</a:lnRef>
						<a:fillRef idx="1">
							<a:schemeClr val="accent1"/>
						</a:fillRef>
						<a:effectRef idx="0">
							<a:schemeClr val="accent1"/>
						</a:effectRef>
						<a:fontRef idx="minor">
							<a:schemeClr val="lt1"/>
						</a:fontRef>
					</p:style>
					<p:txBody>
						<a:bodyPr lIns="72000" tIns="72000" rIns="72000" bIns="72000" rtlCol="0" anchor="ctr"/>
						<a:lstStyle/>
						<a:p>
							<a:pPr marL="0" marR="0" lvl="0" indent="0" algn="ctr" defTabSz="914400" rtl="0" eaLnBrk="1" fontAlgn="auto" latinLnBrk="0" hangingPunct="1">
								<a:lnSpc>
									<a:spcPct val="100%"/>
								</a:lnSpc>
								<a:spcBef>
									<a:spcPts val="0"/>
								</a:spcBef>
								<a:spcAft>
									<a:spcPts val="0"/>
								</a:spcAft>
								<a:buClrTx/>
								<a:buSzTx/>
								<a:buFontTx/>
								<a:buNone/>
								<a:tabLst/>
								<a:defRPr/>
							</a:pPr>
							<a:endParaRPr kumimoji="0" lang="en-US" sz="1600" b="0" i="0" u="none" strike="noStrike" kern="1200" cap="none" spc="0" normalizeH="0" baseline="0%" noProof="0" err="1">
								<a:ln>
									<a:noFill/>
								</a:ln>
								<a:solidFill>
									<a:srgbClr val="000000"/>
								</a:solidFill>
								<a:effectLst/>
								<a:uLnTx/>
								<a:uFillTx/>
								<a:latin typeface="Arial"/>
								<a:ea typeface="+mn-ea"/>
								<a:cs typeface="+mn-cs"/>
							</a:endParaRPr>
						</a:p>
					</p:txBody>
				</p:sp>
				<p:sp>
					<p:nvSpPr>
						<p:cNvPr id="238" name="Textfeld 237"/>
						<p:cNvSpPr txBox="1"/>
						<p:nvPr/>
					</p:nvSpPr>
					<p:spPr>
						<a:xfrm>
							<a:off x="6288399" y="6220624"/>
							<a:ext cx="670818" cy="283906"/>
						</a:xfrm>
						<a:prstGeom prst="rect">
							<a:avLst/>
						</a:prstGeom>
						<a:noFill/>
					</p:spPr>
					<p:txBody>
						<a:bodyPr wrap="square" lIns="72000" tIns="72000" rIns="72000" bIns="72000" rtlCol="0">
							<a:spAutoFit/>
						</a:bodyPr>
						<a:lstStyle/>
						<a:p>
							<a:pPr marL="0" marR="0" lvl="0" indent="0" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" fontAlgn="auto" latinLnBrk="0" hangingPunct="1">
								<a:lnSpc>
									<a:spcPct val="100%"/>
								</a:lnSpc>
								<a:spcBef>
									<a:spcPts val="0"/>
								</a:spcBef>
								<a:spcAft>
									<a:spcPts val="0"/>
								</a:spcAft>
								<a:buClrTx/>
								<a:buSzTx/>
								<a:buFontTx/>
								<a:buNone/>
								<a:tabLst/>
								<a:defRPr/>
							</a:pPr>
							<a:r>
								<a:rPr kumimoji="0" lang="en-US" sz="900" b="0" i="0" u="none" strike="noStrike" kern="1200" cap="none" spc="0" normalizeH="0" baseline="0%" noProof="0" dirty="0">
									<a:ln>
										<a:noFill/>
									</a:ln>
									<a:solidFill>
										<a:prstClr val="black"/>
									</a:solidFill>
									<a:effectLst/>
									<a:uLnTx/>
									<a:uFillTx/>
									<a:latin typeface="Arial"/>
									<a:ea typeface="+mn-ea"/>
									<a:cs typeface="+mn-cs"/>
								</a:rPr>
								<a:t>OI in pcs.</a:t>
							</a:r>
						</a:p>
					</p:txBody>
				</p:sp>
				<p:sp>
					<p:nvSpPr>
						<p:cNvPr id="248" name="Rechteck 247"/>
						<p:cNvSpPr/>
						<p:nvPr/>
					</p:nvSpPr>
					<p:spPr>
						<a:xfrm>
							<a:off x="6924214" y="6317911"/>
							<a:ext cx="288000" cy="108000"/>
						</a:xfrm>
						<a:prstGeom prst="rect">
							<a:avLst/>
						</a:prstGeom>
						<a:solidFill>
							<a:srgbClr val="FF0000"/>
						</a:solidFill>
						<a:ln w="6350">
							<a:solidFill>
								<a:schemeClr val="accent6"/>
							</a:solidFill>
						</a:ln>
					</p:spPr>
					<p:style>
						<a:lnRef idx="2">
							<a:schemeClr val="accent1">
								<a:shade val="50%"/>
							</a:schemeClr>
						</a:lnRef>
						<a:fillRef idx="1">
							<a:schemeClr val="accent1"/>
						</a:fillRef>
						<a:effectRef idx="0">
							<a:schemeClr val="accent1"/>
						</a:effectRef>
						<a:fontRef idx="minor">
							<a:schemeClr val="lt1"/>
						</a:fontRef>
					</p:style>
					<p:txBody>
						<a:bodyPr lIns="72000" tIns="72000" rIns="72000" bIns="72000" rtlCol="0" anchor="ctr"/>
						<a:lstStyle/>
						<a:p>
							<a:pPr marL="0" marR="0" lvl="0" indent="0" algn="ctr" defTabSz="914400" rtl="0" eaLnBrk="1" fontAlgn="auto" latinLnBrk="0" hangingPunct="1">
								<a:lnSpc>
									<a:spcPct val="100%"/>
								</a:lnSpc>
								<a:spcBef>
									<a:spcPts val="0"/>
								</a:spcBef>
								<a:spcAft>
									<a:spcPts val="0"/>
								</a:spcAft>
								<a:buClrTx/>
								<a:buSzTx/>
								<a:buFontTx/>
								<a:buNone/>
								<a:tabLst/>
								<a:defRPr/>
							</a:pPr>
							<a:r>
								<a:rPr kumimoji="0" lang="en-US" sz="500" b="0" i="0" u="none" strike="noStrike" kern="1200" cap="none" spc="0" normalizeH="0" baseline="0%" noProof="0" dirty="0">
									<a:ln>
										<a:noFill/>
									</a:ln>
									<a:solidFill>
										<a:srgbClr val="FFFFFF"/>
									</a:solidFill>
									<a:effectLst/>
									<a:uLnTx/>
									<a:uFillTx/>
									<a:latin typeface="Arial"/>
									<a:ea typeface="+mn-ea"/>
									<a:cs typeface="+mn-cs"/>
								</a:rPr>
								<a:t>New</a:t>
							</a:r>
						</a:p>
					</p:txBody>
				</p:sp>
				<p:sp>
					<p:nvSpPr>
						<p:cNvPr id="278" name="Textbox"/>
						<p:cNvSpPr txBox="1">
							<a:spLocks/>
						</p:cNvSpPr>
						<p:nvPr/>
					</p:nvSpPr>
					<p:spPr>
						<a:xfrm>
							<a:off x="5189149" y="6076963"/>
							<a:ext cx="2880000" cy="246322"/>
						</a:xfrm>
						<a:prstGeom prst="rect">
							<a:avLst/>
						</a:prstGeom>
					</p:spPr>
					<p:txBody>
						<a:bodyPr vert="horz" lIns="0" tIns="0" rIns="0" bIns="0" rtlCol="0">
							<a:noAutofit/>
						</a:bodyPr>
						<a:lstStyle>
							<a:lvl1pPr marL="0" indent="0" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1">
								<a:lnSpc>
									<a:spcPct val="100%"/>
								</a:lnSpc>
								<a:spcBef>
									<a:spcPts val="600"/>
								</a:spcBef>
								<a:buClrTx/>
								<a:buFontTx/>
								<a:buNone/>
								<a:defRPr sz="1400" kern="1200">
									<a:solidFill>
										<a:schemeClr val="tx1"/>
									</a:solidFill>
									<a:latin typeface="+mn-lt"/>
									<a:ea typeface="+mn-ea"/>
									<a:cs typeface="+mn-cs"/>
								</a:defRPr>
							</a:lvl1pPr>
							<a:lvl2pPr marL="216000" indent="-216000" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1">
								<a:lnSpc>
									<a:spcPct val="100%"/>
								</a:lnSpc>
								<a:spcBef>
									<a:spcPts val="600"/>
								</a:spcBef>
								<a:buClrTx/>
								<a:buFont typeface="Symbol" panose="05050102010706020507" pitchFamily="18" charset="2"/>
								<a:buChar char="-"/>
								<a:defRPr sz="1400" kern="1200">
									<a:solidFill>
										<a:schemeClr val="tx1"/>
									</a:solidFill>
									<a:latin typeface="+mn-lt"/>
									<a:ea typeface="+mn-ea"/>
									<a:cs typeface="+mn-cs"/>
								</a:defRPr>
							</a:lvl2pPr>
							<a:lvl3pPr marL="432000" indent="-216000" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1">
								<a:lnSpc>
									<a:spcPct val="100%"/>
								</a:lnSpc>
								<a:spcBef>
									<a:spcPts val="600"/>
								</a:spcBef>
								<a:buClrTx/>
								<a:buFont typeface="Symbol" panose="05050102010706020507" pitchFamily="18" charset="2"/>
								<a:buChar char="-"/>
								<a:defRPr sz="1400" kern="1200">
									<a:solidFill>
										<a:schemeClr val="tx1"/>
									</a:solidFill>
									<a:latin typeface="+mn-lt"/>
									<a:ea typeface="+mn-ea"/>
									<a:cs typeface="+mn-cs"/>
								</a:defRPr>
							</a:lvl3pPr>
							<a:lvl4pPr marL="648000" indent="-216000" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1">
								<a:lnSpc>
									<a:spcPct val="100%"/>
								</a:lnSpc>
								<a:spcBef>
									<a:spcPts val="600"/>
								</a:spcBef>
								<a:buClrTx/>
								<a:buFont typeface="Symbol" panose="05050102010706020507" pitchFamily="18" charset="2"/>
								<a:buChar char="-"/>
								<a:defRPr sz="1400" kern="1200">
									<a:solidFill>
										<a:schemeClr val="tx1"/>
									</a:solidFill>
									<a:latin typeface="+mn-lt"/>
									<a:ea typeface="+mn-ea"/>
									<a:cs typeface="+mn-cs"/>
								</a:defRPr>
							</a:lvl4pPr>
							<a:lvl5pPr marL="864000" indent="-216000" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1">
								<a:lnSpc>
									<a:spcPct val="100%"/>
								</a:lnSpc>
								<a:spcBef>
									<a:spcPts val="600"/>
								</a:spcBef>
								<a:buClrTx/>
								<a:buFont typeface="Symbol" panose="05050102010706020507" pitchFamily="18" charset="2"/>
								<a:buChar char="-"/>
								<a:defRPr sz="1400" kern="1200">
									<a:solidFill>
										<a:schemeClr val="tx1"/>
									</a:solidFill>
									<a:latin typeface="+mn-lt"/>
									<a:ea typeface="+mn-ea"/>
									<a:cs typeface="+mn-cs"/>
								</a:defRPr>
							</a:lvl5pPr>
							<a:lvl6pPr marL="2514600" indent="-228600" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1">
								<a:lnSpc>
									<a:spcPct val="90%"/>
								</a:lnSpc>
								<a:spcBef>
									<a:spcPts val="500"/>
								</a:spcBef>
								<a:buFont typeface="Arial" panose="020B0604020202020204" pitchFamily="34" charset="0"/>
								<a:buChar char="???"/>
								<a:defRPr sz="1800" kern="1200">
									<a:solidFill>
										<a:schemeClr val="tx1"/>
									</a:solidFill>
									<a:latin typeface="+mn-lt"/>
									<a:ea typeface="+mn-ea"/>
									<a:cs typeface="+mn-cs"/>
								</a:defRPr>
							</a:lvl6pPr>
							<a:lvl7pPr marL="2971800" indent="-228600" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1">
								<a:lnSpc>
									<a:spcPct val="90%"/>
								</a:lnSpc>
								<a:spcBef>
									<a:spcPts val="500"/>
								</a:spcBef>
								<a:buFont typeface="Arial" panose="020B0604020202020204" pitchFamily="34" charset="0"/>
								<a:buChar char="???"/>
								<a:defRPr sz="1800" kern="1200">
									<a:solidFill>
										<a:schemeClr val="tx1"/>
									</a:solidFill>
									<a:latin typeface="+mn-lt"/>
									<a:ea typeface="+mn-ea"/>
									<a:cs typeface="+mn-cs"/>
								</a:defRPr>
							</a:lvl7pPr>
							<a:lvl8pPr marL="3429000" indent="-228600" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1">
								<a:lnSpc>
									<a:spcPct val="90%"/>
								</a:lnSpc>
								<a:spcBef>
									<a:spcPts val="500"/>
								</a:spcBef>
								<a:buFont typeface="Arial" panose="020B0604020202020204" pitchFamily="34" charset="0"/>
								<a:buChar char="???"/>
								<a:defRPr sz="1800" kern="1200">
									<a:solidFill>
										<a:schemeClr val="tx1"/>
									</a:solidFill>
									<a:latin typeface="+mn-lt"/>
									<a:ea typeface="+mn-ea"/>
									<a:cs typeface="+mn-cs"/>
								</a:defRPr>
							</a:lvl8pPr>
							<a:lvl9pPr marL="3886200" indent="-228600" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1">
								<a:lnSpc>
									<a:spcPct val="90%"/>
								</a:lnSpc>
								<a:spcBef>
									<a:spcPts val="500"/>
								</a:spcBef>
								<a:buFont typeface="Arial" panose="020B0604020202020204" pitchFamily="34" charset="0"/>
								<a:buChar char="???"/>
								<a:defRPr sz="1800" kern="1200">
									<a:solidFill>
										<a:schemeClr val="tx1"/>
									</a:solidFill>
									<a:latin typeface="+mn-lt"/>
									<a:ea typeface="+mn-ea"/>
									<a:cs typeface="+mn-cs"/>
								</a:defRPr>
							</a:lvl9pPr>
						</a:lstStyle>
						<a:p>
							<a:pPr marL="0" lvl="1" indent="0">
								<a:buNone/>
							</a:pPr>
							<a:r>
								<a:rPr lang="en-US" sz="900" dirty="0"/>
								<a:t>Source: QlikView Global Report</a:t>
							</a:r>
						</a:p>
					</p:txBody>
				</p:sp>
			</p:grpSp>
			<xsl:comment>
				<xsl:text>
			
			*** End added portion of shape tree ***
			
</xsl:text>
			</xsl:comment>
		</xsl:copy>
	</xsl:template>

	<xsl:template match="@* | node()">
		<xsl:copy>
			<xsl:apply-templates select="@* | node()" />
		</xsl:copy>
	</xsl:template>
	
	<xsl:template match="/">
		<xsl:comment><xsl:text>*** This file has been modified ***</xsl:text></xsl:comment>
		<xsl:apply-templates select="@* | node()" />
	</xsl:template>
	
</xsl:stylesheet>
"@
#
# -----------------------------------------------------------------------------------------------
#
[Double] $script:decoWidth = 600.0
[Double] $script:decoHeight = 400.0
#
[System.Drawing.SolidBrush] $script:decoBrush = New-Object -TypeName System.Drawing.SolidBrush $([System.Drawing.Color]::FromName('White'))
[System.Drawing.SolidBrush] $script:decoPen = New-Object -TypeName System.Drawing.SolidBrush $([System.Drawing.Color]::FromName('Black'))
[System.Drawing.Font] $script:decofontNormal = New-Object -Typename System.Drawing.Font "Courier New", $(10 * $script:decoWidth * 0.006)
[System.Drawing.FontStyle] $script:decostyleBold = [System.Drawing.Fontstyle]::Bold
[System.Drawing.Font] $script:decofontBold = New-Object -Typename System.Drawing.Font $script:decofontNormal, $script:decostyleBold
#
function local:scaleImage ([String] $decoInFile, [String] $decoOutFile) {
	#
	[System.Drawing.Bitmap] $local:scaledBitmap = New-Object -TypeName System.Drawing.Bitmap @([Convert]::ToInt32($script:decoWidth), [Convert]::ToInt32($script:decoHeight))
	#
	[System.Drawing.Graphics] $local:graph = [System.Drawing.Graphics]::FromImage($private:scaledBitmap)
	#
	$local:graph.InterpolationMode = [System.Drawing.Drawing2D.InterpolationMode]::High
	$local:graph.CompositingQuality = [System.Drawing.Drawing2D.CompositingQuality]::HighQuality
	$local:graph.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
	$local:graph.FillRectangle($script:decoBrush, $(New-Object -TypeName System.Drawing.Rectangle @(0, 0, [Convert]::ToInt32($script:decoWidth), [Convert]::ToInt32($script:decoHeight))))
	#
	if (Test-Path -Path $decoInFile) {
		#
		[System.Drawing.Bitmap] $private:rawImg = [System.Drawing.Bitmap]::FromFile($decoInFile, $true)
		#
		[Double] $private:scale = [Math]::Min($script:decoWidth / $($private:rawImg.Width), $script:decoHeight / $($private:rawImg.Height))
		#
		[Int] $private:scaleWitdh = [Convert]::ToInt32($($private:rawImg.Width) * $private:scale)
		[Int] $private:scaleHeight = [Convert]::ToInt32($($private:rawImg.Height) * $private:scale)
		#
        Write-Verbose -Message "Scaling image from $decoInFile to $([Int]$($private:scale * 100))%"
        #
		$local:graph.DrawImage($private:rawImg, $(New-Object -TypeName System.Drawing.Rectangle @([Convert]::ToInt32(0.1 * ($script:decoWidth - $private:scaleWitdh)) , [Convert]::ToInt32(0.5 * ($script:decoHeight -  $private:scaleHeight)), $private:scaleWitdh, $private:scaleHeight)))
		#
		$private:rawImg.Dispose()
		#
	} else {
		#
		Write-Verbose -Message "Image file $decoInFile not found."
		#
		[System.Drawing.RectangleF] $private:decorect = New-Object -TypeName System.Drawing.RectangleF $($script:decoWidth * 0.01), $($script:decoWidth * 0.1), $($script:decoWidth*0.99), $($script:decoWidth * 0.9)
		#
		$local:graph.DrawString($decoInFile, $script:decoFontBold, $script:decoPen, $script:decoRect)
	}
	#
	$local:scaledBitmap.Save($decoOutFile, [System.Drawing.Imaging.ImageFormat]::Jpeg)
	#
	$local:graph.Dispose()
	#
	$local:scaledBitmap.Dispose()
	#
}
#
# -----------------------------------------------------------------------------------------------
#
function script:msxml([Object] $xsl, [Object] $xml, [String] $out = '', [System.Collections.Hashtable] $param = @{}, [Boolean] $verbose = $false, [Boolean] $debug = $false) {
	#
	# -----------------------------------------------------------------------------------------------
	#
	#	Werte der Variablen $VerbosePreference und $DebugPreference sichern
	#	und entsprechend den Parametern $verbose und $debug neu setzen.
	#
	[System.Management.Automation.ActionPreference] $script:saveVerbosePref = $VerbosePreference
	[System.Management.Automation.ActionPreference] $script:saveDebugPref = $DebugPreference
	#
	if($verbose -eq $true -or  $debug -eq $true) {
		#
		Write-Verbose @"
	msxml(...): switched to verbose mode.
"@
		#
		$VerbosePreference = [System.Management.Automation.ActionPreference]::Continue
		#
		if($debug -eq $true) {
			Write-Verbose @"
	msxml(...): switched to debug mode.
"@
			$DebugPreference = [System.Management.Automation.ActionPreference]::Inquire
		} else {
			$DebugPreference = [System.Management.Automation.ActionPreference]::Continue
		}
	} else {
		$VerbosePreference = [System.Management.Automation.ActionPreference]::SilentlyContinue
		$DebugPreference = [System.Management.Automation.ActionPreference]::SilentlyContinue
	}
	#
	# -----------------------------------------------------------------------------------------------
	#
	#
	[System.Xml.XmlUrlResolver] $local:xres = New-Object -typeName System.Xml.XmlUrlResolver
	[System.Xml.Xsl.XsltSettings] $local:xset = New-Object -typeName System.Xml.Xsl.XsltSettings
	#
	$null = $xset.set_EnableDocumentFunction($true)
	$null = $xset.set_EnableScript($true)
	#
	[System.Xml.Xsl.XslCompiledTransform] $script:xslt = New-Object -typeName System.Xml.Xsl.XslCompiledTransform
	[System.Xml.Xsl.XsltArgumentList] $script:xal = New-Object -typename System.Xml.Xsl.XsltArgumentList
	#
	if($verbose -eq $true) {
		$script:xal.AddParam("verbose", "", 1)
	} else {
		$script:xal.AddParam("verbose", "", 0)
	}
	if($debug -eq $true) {
		$script:xal.AddParam("debug", "", 1)
	} else {
		$script:xal.AddParam("debug", "", 0)
	}
	#
	if($null -ne $param) {
		[String[]] $local:val = $null
		foreach($val in $param.get_Keys()) {
			Write-Debug @"
[DEBUG] msxml::msxml(...): Adding XSLT parameter $val = $($param[$val])
"@
			$script:xal.AddParam($val, "", $($param[$val]))
		}
	}
	#
	if($xsl.GetType().Name -eq 'XmlDocument') {
		#
		$xslt.Load([System.Xml.XmlReader]::Create($(New-Object -typeName System.IO.StringReader -argumentList $xsl.PsBase.InnerXml)), $local:xset, $local:xres)
		#
	} else {
		#
		[String] $local:fullXsl = $(Get-Item $xsl | Select-Object FullPath -ExpandProperty FullName)
		#
		if ([System.IO.File]::Exists($local:fullXsl) -eq $true) {
		#
		$xslt.Load($local:fullXsl, $local:xset, $local:xres)
		#
		} else {
			#
		Write-Verbose @"
[FATAL] msxml::msxml(...): Transformation failed (1).
"@
			#
		}
		#
	}
	#
	if ($out -ne '') {
		#
		[System.IO.FileStream] $local:fsm = New-Object System.IO.FileStream -ArgumentList @($out, 2)
		#
		if($xml.GetType().Name -eq 'XmlDocument') {
			#
			$xslt.Transform([System.Xml.XmlReader]::Create($(New-Object -typeName System.IO.StringReader -argumentList $xml.PsBase.InnerXml)), $script:xal, $local:fsm)
			#
		} else {
			#
			[String] $local:fullXml = $(Get-Item $xml | Select-Object FullPath -ExpandProperty FullName)
			#
			if ([System.IO.File]::Exists($local:fullXml) -eq $true) {
			#
			$xslt.Transform($local:fullXml, $script:xal, $local:fsm)
			#
			} else {
				Write-Verbose @"
[FATAL] msxml::msxml(...): Transformation failed (2).
"@
			}
			#
		}
		#
		$local:fsm.Close()
		$local:fsm.Dispose()
		#
		# -----------------------------------------------------------------------------------------------
		#
		#	Werte vor Start des Skripts wiederherstellen
		#
		$VerbosePreference = $script:saveVerbosePref
		$DebugPreference = $script:saveDebugPref
		#
		# -----------------------------------------------------------------------------------------------
		#
		return $null
		#
	} else {
		#	
		[System.IO.StringWriter] $local:osw = New-Object System.IO.StringWriter
		#
		if($xml.GetType().Name -eq 'XmlDocument') {
			#
			$xslt.Transform([System.Xml.XmlReader]::Create($(New-Object -typeName System.IO.StringReader -argumentList $xml.PsBase.InnerXml)), $script:xal, $local:osw)
			#
		} else {
			#
			[String] $local:fullXml = $(Get-Item $xml | Select-Object FullPath -ExpandProperty FullName)
			#
			if ([System.IO.File]::Exists($local:fullXml) -eq $true) {
			#
			$xslt.Transform($local:fullXml, $script:xal, $local:osw)
			#
			} else {
				#
				Write-Verbose @"
[FATAL] msxml::msxml(...): Transformation failed (3).
"@
			}
			#
		}
		$local:osw.close()
		[String] $local:res = $local:osw.ToString()
		$local:osw.Dispose()
		#
		# -----------------------------------------------------------------------------------------------
		#
		#	Werte vor Start des Skripts wiederherstellen
		#
		$VerbosePreference = $script:saveVerbosePref
		$DebugPreference = $script:saveDebugPref
		#
		# -----------------------------------------------------------------------------------------------
		#
		return $local:res
	}
}
#
# -----------------------------------------------------------------------------------------------
#
function script:OpenFileDialog([String] $title = "Oups.", [String] $type = "all", [String] $defpath = "$pwd") {
	#
	[System.Windows.Forms.OpenFileDialog] $openFileDialog1 = new-object -typeName System.Windows.Forms.OpenFileDialog
	#
	[System.Collections.Hashtable] $local:filters = @{'pptx'='Powerpoint Slideshow (*.pptx)|*.pptx'; 'pdf'='Portable Data Format (*.pdf)|*.pdf'; 'fo'='Formatting Objects (*.fo)|*.fo';'xmlcsv'='Excel 2003 XML or Excel CSV file (*.xml)|*.xml;*csv';'csv'='Comma separated values spread sheet (*.csv)|*.csv'; 'htm'='HTML Source (*.htm, *.html)|*.htm;*.html'; 'xslt'= 'XSLT Stylesheet (*.xsl, *.xslt)|*.xsl;*.xslt'; 'i6z'='IUCLID (*.i6z)|*.i6z'; 'all'='All files (*.*)|*.*'}
	#
	[String] $local:res = $null
	#
	$openFileDialog1.Reset()
	#
	if([System.IO.Directory]::Exists($defpath)) {
		Write-Debug @"
[Debug] framework::OpenFileDialog: Default path set to $defpath
"@
		$openFileDialog1.set_InitialDirectory($defpath)
	} else {
		Write-Debug @"
[Debug] framework::OpenFileDialog: Default path not set.
"@
	}
	$openFileDialog1.set_Filter($filters[$type])
	$openFileDialog1.set_FilterIndex(2)
	$openFileDialog1.set_Title($title)
	$openFileDialog1.set_AddExtension($true)
	$openFileDialog1.set_AutoUpgradeEnabled($true)
	$openFileDialog1.set_CheckFileExists($true)
	$openFileDialog1.set_CheckPathExists($true)
	$openFileDialog1.set_ShowHelp($false)
	$openFileDialog1.set_RestoreDirectory($true)
	$openFileDialog1.set_Multiselect($false)
	$openFileDialog1.set_ShowReadOnly($false)
	#
	if($openFileDialog1.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK)
	{
		$res = $openFileDialog1.get_FileName()
        } else {
		$res = $null
	}
	#
	$openFileDialog1.Dispose()
	#
	return $res
}
#
# -----------------------------------------------------------------------------------------------
#
#	Werte der Variablen $VerbosePreference und $DebugPreference sichern
#	und entsprechend den Parametern $verbose und $debug neu setzen.
#
[System.Management.Automation.ActionPreference] $script:saveVerbosePref = $VerbosePreference
[System.Management.Automation.ActionPreference] $script:saveDebugPref = $DebugPreference
#
if($verbose -eq $true -or  $debug -eq $true) {
	#
	Write-Verbose @"
framework.ps1::main(...): switched to verbose mode.
"@
	#
	$VerbosePreference = [System.Management.Automation.ActionPreference]::Continue
	#
	if($debug -eq $true) {
		Write-Verbose @"
framework.ps1::main(...): switched to debug mode.
"@
		$DebugPreference = [System.Management.Automation.ActionPreference]::Inquire
		#
	} else {
		#
		$DebugPreference = [System.Management.Automation.ActionPreference]::SilentlyContinue
		#
	}
} else {
	#
    $VerbosePreference = [System.Management.Automation.ActionPreference]::SilentlyContinue
    $DebugPreference = [System.Management.Automation.ActionPreference]::SilentlyContinue
	#
}
#
# -----------------------------------------------------------------------------------------------
#
if (($verbose -or $debug) -ne $true) { Clear-Host }
#
[System.Windows.Forms.FolderBrowserDialog] $private:folderBrowserDialog = New-Object -TypeName System.Windows.Forms.FolderBrowserDialog 
#
$private:folderBrowserDialog.set_SelectedPath($DataDir)
$private:folderBrowserDialog.set_ShowNewFolderButton($false)
$private:folderBrowserDialog.set_Description("Wo sollen die Bilder gesucht werden?")
#
if ($private:folderBrowserDialog.ShowDialog() -eq 'OK') {
	#
	[String] $private:folder = $private:folderBrowserDialog.get_SelectedPath()
    #
	[String] $local:slideshow = $(. script:OpenFileDialog -title "Powerpoint Datei zum Bearbeiten" -defpath $DataDir -type 'pptx')
	#
	if (Test-Path -Path $local:slideshow) {
		#
        Set-Location -Path $DataDir
        #    
        New-Item -ItemType Directory -Path $DataDir -Name "ppt" -Force | Write-Verbose
        New-Item -ItemType Directory -Path $DataDir\ppt -Name "slides" -Force | Write-Verbose
        New-Item -ItemType Directory -Path $DataDir\ppt\slides -Name "_rels" -Force | Write-Verbose
        #
		. 'C:\Program Files\7-Zip\7z.exe' l -y $local:slideshow |  Write-Debug
		#
		. 'C:\Program Files\7-Zip\7z.exe' x -y $local:slideshow "ppt\slides\slide1.xml" |  Write-Verbose
		. 'C:\Program Files\7-Zip\7z.exe' x -y $local:slideshow "ppt\slides\_rels\slide1.xml.rels" |  Write-Verbose
		#
		[Int] $local:imgOff = [Int]$($($(. 'C:\Program Files\7-Zip\7z.exe' l -y $local:slideshow) -match "ppt\\media") -replace ".*ppt\\media\\image([1-9][0-9]*)\.[A-Za-z]{3,4}",'$1' | Sort-Object | Select-Object -Last 1) + 1
		#
        [Int] $local:relOff = ($([Xml]$(Get-Content "$DataDir\ppt\slides\_rels\slide1.xml.rels")).Relationships.Relationship.Id -replace '^rId([1-9][0-9]?)','$1' | ForEach-Object -Process { [Int]$_} | Sort-Object | Select-Object -Last 1) + 1
        #
        Write-Verbose "First image index is: ........ $([Convert]::ToString($local:imgOff))"
        Write-Verbose "First relation index is: ..... $([Convert]::ToString($local:relOff))"
		#
		[String] $local:dataFile = $(. script:OpenFileDialog -title "Daten aus CSV UTF-8 Tabelle" -defpath $DataDir -type 'csv')
		#
		if (Test-Path -Path $local:dataFile) {
			#
			New-Item -Type Directory -Path $DataDir\ppt -Name "media" -Force | Write-Verbose
			#
			[Xml] $local:tagData = $(Import-Csv -Encoding utf8 -Delimiter ";" -Path $local:dataFile | ConvertTo-Xml)
			#
            [Int] $local:column = -1
            #
            foreach($pos in 0..$($tagData.Objects.Object[1].Property.Length - 1)) {
                #
                if($tagData.Objects.Object[1].Property[$pos].Name -eq 'image') {
                    #
                    $local:column = $pos
                    #
                    break
                    #
                }
                #
            }
            #
            if ($local:column -ne -1) {
                #
				[Int] $local:imgIdx = $local:imgOff
				[Int] $local:count = 0
				#
				foreach($item in $local:tagData.Objects.Object) {
					#
					Write-Progress -Id  345 -Activity "Rescaling images ..." -PercentComplete $([int]($local:count++ / $local:tagData.Objects.Object.Length * 100))
					#
					. local:scaleImage -decoInFile "$private:folder\$($item.Property[$private:column].'#text')" -decoOutFile "$DataDir\ppt\media\image$local:imgIdx.jpeg"
					#
					$local:imgIdx++
					#
				}
				#
				[String] $private:tmpText8 = [System.IO.Path]::GetTempFileName()
				[String] $private:tmpSide8 = [System.IO.Path]::GetTempFileName()
				#
				Out-File -FilePath $private:tmpSide8 -InputObject $local:tagData.PsBase.InnerXml -Encoding utf8
				#
				#	$(Import-Csv -Encoding utf8 -Delimiter ";" -Path $local:load | ConvertTo-Xml)
				#
				Write-Verbose "Transformation der ersten Folie in $local:slideshow mit dem Stylesheet $local:stylesheet"
				#
				[String] $private:tmpText8 = [System.IO.Path]::GetTempFileName()
				#
				[String] $private:message = $(. msxml -xsl $transform -xml "$DataDir\ppt\slides\slide1.xml" -out "$private:tmpText8" -param @{'Decorate.FirstRelationIndex'="$local:relOff";"Decorate.TagDataLoadFile"="$private:tmpSide8" })
				#
                if ($private:message.Trim().Length -gt 0) {
                    #
                    Write-Host "[XSLT] $private:message"
                    #
                }
                #
				Move-Item -LiteralPath $private:tmpText8 -Destination "$DataDir\ppt\slides\slide1.xml" -Force 
				#
				$private:message = $(. msxml -xsl $transform -xml "$DataDir\ppt\slides\_rels\slide1.xml.rels" -out "$private:tmpText8" -param @{'Decorate.FirstImageIndex'="$local:imgOff";"Decorate.TagDataLoadFile"="$private:tmpSide8"})
				#
                if ($private:message.Trim().Length -gt 0) {
                    #
                    Write-Host "[XSLT] $private:message"
                    #
                }
                #
				Move-Item -LiteralPath $private:tmpText8 -Destination "$DataDir\ppt\slides\_rels\slide1.xml.rels" -Force
				#
				if(Test-Path -Path $private:tmpSide8) { Remove-Item -Path $private:tmpSide8 }
				#
			} else {
                #
                Write-Verbose -Message "[ERROR] Data file has no column named 'image' - terminating."
                #
            }
			#
		}
		#
		. 'C:\Program Files\7-Zip\7z.exe' d -y $local:slideshow "ppt\slides\slide1.xml" | Write-Verbose
		. 'C:\Program Files\7-Zip\7z.exe' d -y $local:slideshow "ppt\slides\_rels\slide1.xml.rels" | Write-Verbose
		#
		. 'C:\Program Files\7-Zip\7z.exe' a -sdel -y $local:slideshow "ppt\slides\slide1.xml" | Write-Verbose
		. 'C:\Program Files\7-Zip\7z.exe' a -sdel -y $local:slideshow "ppt\slides\_rels\slide1.xml.rels" | Write-Verbose
		. 'C:\Program Files\7-Zip\7z.exe' a -sdel -y $local:slideshow "ppt\media\image*.jpeg" | Write-Verbose
		#
        Remove-Item -Path "$DataDir\ppt" -Recurse -Force | Write-Verbose
        #
	}
	#
}
#
#
# -----------------------------------------------------------------------------------------------
#
#	Werte vor Start des Skripts wiederherstellen
#
$VerbosePreference = $script:saveVerbosePref
$DebugPreference = $script:saveDebugPref
#
# -----------------------------------------------------------------------------------------------
#