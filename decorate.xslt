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
	
	<xsl:variable name="Decorate.TagData" select="document($Decorate.TagDataLoadFile)/Objects/Object" />
	<!--

	-->
	<xsl:output encoding="UTF-8" method="xml" standalone="yes" indent="yes" />
	<!--
		CustomData und Extensions werden nicht mitgenommen
	-->
	<xsl:template match="a:extLst|p:custDataLst" />

	<xsl:template name="extraRels" xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
		<xsl:param name="startId" />
		<xsl:for-each select="$Decorate.TagData">
			<Relationship Id="{concat ('rId', ($startId + position () - 1))}" Type="http://purl.oclc.org/ooxml/officeDocument/relationships/image" Target="{concat ('../media/image', ($Decorate.FirstImageIndex + position() - 1), '.jpeg')}"/>
		</xsl:for-each>
	</xsl:template>

	<xsl:template name="recRels">
		<xsl:param name="rels" />
		<xsl:param name="maxId" select="0" />
		<xsl:choose>
			<xsl:when test="$rels">
				<xsl:apply-templates select="$rels [1]" />
				<xsl:call-template name="recRels">
					<xsl:with-param name="rels" select="$rels [position() > 1]" />
					<xsl:with-param name="maxId">
						<xsl:variable name="rId" select="number (substring-after ($rels [1]/@Id, 'rId'))" />
						<xsl:choose>
							<xsl:when test="$rId > $maxId">
								<xsl:value-of select="$rId" />
							</xsl:when>
							<xsl:otherwise>
								<xsl:value-of select="$maxId" />
							</xsl:otherwise>
						</xsl:choose>
					</xsl:with-param>
				</xsl:call-template>
			</xsl:when>
			<xsl:otherwise>
				<xsl:comment>
					<xsl:text>
				
			*** Ende der mitgenommenen rels, max. rId=</xsl:text><xsl:value-of select="$maxId" /><xsl:text> ***
			
	</xsl:text>
				</xsl:comment>
				<xsl:call-template name="extraRels">
					<xsl:with-param name="startId" select="$maxId + 1" />
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
				<p:cNvPr id="83" name="{concat ('Hintergrund an Position x=', $xpos, ' y=', $ypos)}"/>
				<p:cNvSpPr/>
				<p:nvPr/>
			</p:nvSpPr>
			<p:spPr>
				<a:xfrm>
					<a:off x="{279473 + ($xpos - 1) * $xincr}" y="{1049454 + ($ypos -1) * $yincr}"/>
					<a:ext cx="1363980" cy="502920"/>
				</a:xfrm>
				<a:prstGeom prst="rect">
					<a:avLst/>
				</a:prstGeom>
				<a:solidFill>
					<a:srgbClr val="FFFFFF"/>
				</a:solidFill>
				<a:ln w="{$boxlinewidth}">
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
					<p:cNvPr id="12" name="{concat('Symbolfoto an Position x=', $xpos, ' y=', $ypos)}"/>
					<p:cNvPicPr>
						<a:picLocks noChangeAspect="1"/>
					</p:cNvPicPr>
					<p:nvPr/>
				</p:nvPicPr>
				<p:blipFill>
					<a:blip r:embed="{concat ('rId', $imgrel)}"/>
					<a:stretch>
						<a:fillRect/>
					</a:stretch>
				</p:blipFill>
				<p:spPr>
					<a:xfrm>
						<a:off x="{350000 + ($xpos - 1) * $xincr}" y="{1100000 + ($ypos - 1) * $yincr}"/>
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
				<p:cNvPr id="87" name="{concat('Textfeldan an Position x=', $xpos, ' y=', $ypos)}"/>
				<p:cNvSpPr/>
				<p:nvPr/>
			</p:nvSpPr>
			<p:spPr>
				<a:xfrm>
					<a:off x="{1040000 + ($xpos - 1) * $xincr}" y="{1088000 + ($ypos -1) * $yincr}"/>
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
						<a:t><xsl:value-of select="$line1" /></a:t>
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
						<a:t><xsl:value-of select="$line2" /></a:t>
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
						<a:t><xsl:value-of select="$line3" /></a:t>
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
						<a:t><xsl:value-of select="$line4" /></a:t>
					</a:r>
				</a:p>
			</p:txBody>
		</p:sp>
		<!-- Roter Rahmen -->
		<p:sp>
			<p:nvSpPr>
				<p:cNvPr id="33" name="{concat ('Roter Rahmen an Position x=', $xpos, ' y=', $ypos)}"/>
				<p:cNvSpPr/>
				<p:nvPr/>
			</p:nvSpPr>
			<p:spPr>
				<a:xfrm>
					<a:off x="{279473 + ($xpos - 1) * $xincr}" y="{1049454 + ($ypos -1) * $yincr}"/>
					<a:ext cx="1363980" cy="502920"/>
				</a:xfrm>
				<a:prstGeom prst="rect">
					<a:avLst/>
				</a:prstGeom>
				<a:noFill/>
				<a:ln w="{$boxlinewidth}">
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
		<xsl:if test="$check > 0">
			<p:sp>
			<p:nvSpPr>
				<p:cNvPr id="277" name="{concat('Gruener Haken an Position x=', $xpos, ' y=', $ypos)}"/>
				<p:cNvSpPr>
					<a:spLocks noChangeAspect="1"/>
				</p:cNvSpPr>
				<p:nvPr/>
			</p:nvSpPr>
			<p:spPr>
				<a:xfrm>
					<a:off x="{1604666 + ($xpos - 1) * $xincr}" y="{996615 + ($ypos -1) * $yincr}"/>
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
				<p:cNvPr id="145" name="{concat('Gelbes Feld an Position x=', $xpos, ' y=', $ypos)}"/>
				<p:cNvSpPr/>
				<p:nvPr/>
			</p:nvSpPr>
			<p:spPr>
				<a:xfrm>
					<a:off x="{944318 + ($xpos - 1) * $xincr}" y="{966022 + ($ypos - 1) * $yincr}"/>
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
						<a:t><xsl:value-of select="$pieces" /><xsl:text> pcs.</xsl:text></a:t>
					</a:r>
				</a:p>
			</p:txBody>
		</p:sp>
		<!-- Roter Kasten "NEW" -->
		<xsl:if test="$new > 0">
			<p:sp>
				<p:nvSpPr>
					<p:cNvPr id="145" name="{concat('Rotes Feld an Position x=', $xpos, ' y=', $ypos)}"/>
					<p:cNvSpPr/>
					<p:nvPr/>
				</p:nvSpPr>
				<p:spPr>
					<a:xfrm>
						<a:off x="{944318 + ($xpos - 1) * $xincr}" y="{966022 + ($ypos - 1) * $yincr}"/>
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
			<xsl:for-each select="$Decorate.TagData">
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
					<xsl:with-param name="imgrel" select="$Decorate.FirstRelationIndex + position() - 1" />
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
