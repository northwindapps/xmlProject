//
//  styles.swift
//  xmlProject
//
//  Created by yujin on 2020/10/16.
//  Copyright © 2020 yujin. All rights reserved.
//

import Foundation

class Styles{
    
    let styleSheet:String
    let fonts:String
    let font:String
    let fills:String
    let fill:String
    let borders:String
    let border:String
    
    let cellStyleXfs:String
    let cellXfs:String
    let cellStyles:String
    let dxfs:String
    let tableStyles:String
    
    let extLst:String
    
    init() {
        styleSheet="<styleSheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" mc:Ignorable=\"x14ac x16r2 xr\" xmlns:x14ac=\"http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac\" xmlns:x16r2=\"http://schemas.microsoft.com/office/spreadsheetml/2015/02/main\" xmlns:xr=\"http://schemas.microsoft.com/office/spreadsheetml/2014/revision\">"
        
        fonts="<fonts count=\"1\">"
        font="<font><sz val=\"11\"/><color theme=\"1\"/><name val=\"ＭＳ Ｐゴシック\"/><family val=\"2\"/><scheme val=\"minor\"/>"
        
        fills="<fills count=\"2\">"
        
        fill="<fill><patternFill patternType=\"none\"/></fill><fill><patternFill patternType=\"gray125\"/></fill>"
        
        borders="<borders count=\"1\">"
        
        border="<border><left/><right/><top/><bottom/><diagonal/></border>"
        
        cellStyleXfs="<cellStyleXfs count=\"1\"><xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\"/></cellStyleXfs>"
        
        cellXfs="<cellXfs count=\"1\"><xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\"/></cellXfs>"
        
        cellStyles="<cellStyles count=\"1\"><cellStyle name=\"標準\" xfId=\"0\" builtinId=\"0\"/></cellStyles>"
        
        dxfs="<dxfs count=\"0\"/>"
        
        tableStyles="<tableStyles count=\"0\" defaultTableStyle=\"TableStyleMedium2\" defaultPivotStyle=\"PivotStyleMedium9\"/>"
        
        extLst="<extLst><ext uri=\"{EB79DEF2-80B8-43e5-95BD-54CBDDF9020C}\" xmlns:x14=\"http://schemas.microsoft.com/office/spreadsheetml/2009/9/main\"><x14:slicerStyles defaultSlicerStyle=\"SlicerStyleLight1\"/></ext><ext uri=\"{9260A510-F301-46a8-8635-F512D64BE5F5}\" xmlns:x15=\"http://schemas.microsoft.com/office/spreadsheetml/2010/11/main\"><x15:timelineStyles defaultTimelineStyle=\"TimeSlicerStyleLight1\"/></ext></extLst>"
        
    }
    

    //ok static
    func export(){
        var xml = ""
        
        xml.append("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>")
        xml.append(styleSheet)
        xml.append(fonts)
        xml.append(font)
        xml.append("</font>")
        xml.append("</fonts>")
        xml.append(fills)
        xml.append(fill)
        xml.append("</fills>")
        xml.append(borders)
        xml.append(border)
        xml.append("</borders>")
        xml.append(cellStyleXfs)
        xml.append(cellXfs)
        xml.append(cellStyles)
        xml.append(dxfs)
        xml.append(tableStyles)
        xml.append(extLst)
        xml.append("</styleSheet>")
        
        FileManager.default.writeXml(folder: "xl/", filename: "styles.xml", content: xml)
    }
}


//<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
//<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="x14ac x16r2 xr" xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac" xmlns:x16r2="http://schemas.microsoft.com/office/spreadsheetml/2015/02/main" xmlns:xr="http://schemas.microsoft.com/office/spreadsheetml/2014/revision">
//    <fonts count="1">
//    <font>
//        <sz val="11"/>
//        <color theme="1"/>
//        <name val="ＭＳ Ｐゴシック"/>
//        <family val="2"/>
//        <scheme val="minor"/>
//    </font>
//    </fonts>
//    <fills count="2">
//        <fill>
//            <patternFill patternType="none"/>
//        </fill>
//        <fill>
//            <patternFill patternType="gray125"/>
//        </fill>
//    </fills>
//    <borders count="1">
//        <border><left/><right/><top/><bottom/><diagonal/>
//        </border>
//    </borders>
//    <cellStyleXfs count="1">
//        <xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
//    </cellStyleXfs>
//    <cellXfs count="1">
//        <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
//    </cellXfs>
//    <cellStyles count="1">
//        <cellStyle name="標準" xfId="0" builtinId="0"/>
//    </cellStyles>
//    <dxfs count="0"/>
//    <tableStyles count="0" defaultTableStyle="TableStyleMedium2" defaultPivotStyle="PivotStyleMedium9"/>
//    <extLst>
//        <ext uri="{EB79DEF2-80B8-43e5-95BD-54CBDDF9020C}" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main">
//            <x14:slicerStyles defaultSlicerStyle="SlicerStyleLight1"/>
//        </ext>
//        <ext uri="{9260A510-F301-46a8-8635-F512D64BE5F5}" xmlns:x15="http://schemas.microsoft.com/office/spreadsheetml/2010/11/main">
//            <x15:timelineStyles defaultTimelineStyle="TimeSlicerStyleLight1"/>
//        </ext>
//    </extLst>
//</styleSheet>
//


