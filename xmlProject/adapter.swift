//
//  adapter.swift
//  xmlProject
//
//  Created by yujin on 2020/10/19.
//  Copyright © 2020 yujin. All rights reserved.
//

import Foundation

class Adapter{
    var content:[String]
    var location:[String]
    var sheetIdx:[Int]
    var uniqueContent:[String]
    var span:[String]
    let sheetSize:Int
    var rowSpan:[String]
    
    init(imp_content:[String],imp_location:[String],imp_sheetIdx:[Int],imp_sheetSize:Int) {
        content = imp_content//tank,plane
        location = imp_location//A4,B2,A1
        sheetIdx = imp_sheetIdx//1,2,3,2,2
        sheetSize = imp_sheetSize
        uniqueContent = [String]()
        span=[String]()//count = sheetSize
        rowSpan = [String]()//1:2,3:3 count = sheetSize
    }
    
    func getColumnNameNumber(columnName:String)->Int{
        //set span values X:Y
        //Convert column names to number
        //https://stackoverflow.com/questions/667802/what-is-the-algorithm-to-convert-an-excel-column-letter-into-its-number
        let letters = columnName//"AA"
        var sum = 0

        let character: Character = "A"
        let scalars = character.unicodeScalars
        let firstScalar = scalars[scalars.startIndex]
        

        for i in 0..<letters.count {
            
            let characterAti: Character = letters[letters.index(letters.startIndex, offsetBy: i)]
            let scalarsAti = characterAti.unicodeScalars
            let firstScalarAti = scalarsAti[scalarsAti.startIndex].value
            
            
            sum = sum * 26
            sum = sum + (Int(firstScalarAti) - (Int(firstScalar.value) - 1))
        }
        
        print("columnNumber",sum)
        return sum
    }
    
    func initSpanAry(){
        //TODO emptySheet
        for i in 1..<sheetSize+1 {
            var tempSpanValues=[Int]()
            var temprowSpanValues=[Int]()
            for j in 0..<sheetIdx.count {
                if i==sheetIdx[j] {
                    let CN = location[j].letters
                    tempSpanValues.append(getColumnNameNumber(columnName: CN))//AA
                    temprowSpanValues.append(Int.parse(from: location[j])!)//1
                }
            }
            let max = tempSpanValues.max()
            let min = tempSpanValues.min()
            
            if max == nil || min == nil{
                span.append("nil")
                rowSpan.append("nil")
            }else{
                //1:1,1:3
                span.append("spans=\"" + String(min!) + ":" + String(max!) + "\"")
                let max_row = temprowSpanValues.max()
                let min_row = temprowSpanValues.min()
                rowSpan.append(String(min_row!) + ":" + String(max_row!))
            }
        }
        //spans="1:1",rowSpans="1:3"
    }
    
    func initUniqueContentAry(){
        var ary = [String]()
        for i in 0..<content.count {
            if (ary.firstIndex(of:content[i]) == nil) {
                if Double(content[i]) == nil {
                    ary.append(content[i])
                }
            }
        }
        uniqueContent = ary
    }
    
    func createContentArys()->([String],String){
        initSpanAry()
        initUniqueContentAry()
        
        //Create sharedString
        var sharedString_data = "<sst xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" count=\"" + String(content.count) + "\">"
        
        for i in 0..<uniqueContent.count {
            sharedString_data.append("<si><t>")
            sharedString_data.append(uniqueContent[i])
            sharedString_data.append("</t></si>")
        }
        sharedString_data.append("</sst>")
        
        //Creating sheet_data_ary
        //TODO
        var sheet_data_ary = [String]()
        for i in 0..<sheetSize {
            
            let sheet_spans = span[i]
            var sheet_content = [String]()
            var sheet_location = [String]()
            let row_start :Int
            let row_end :Int
            let row_diff :Int
            
            if sheet_spans.contains("nil"){
                sheet_data_ary.append("<sheetData/>")
            }else{
                
            
            row_start = Int(rowSpan[i].components(separatedBy: ":")[0])!
            row_end = Int(rowSpan[i].components(separatedBy: ":")[1])!
            row_diff = row_end - row_start
            
            var sheet_string = ""
            
            //extracting each sheet
            for j in 0..<location.count {
                if i+1 == sheetIdx[j] {//sheet1,sheet2,sheet3...
                    sheet_content.append(content[j])
                    sheet_location.append(location[j])
                }
            }
            
            if row_diff == 0 {
                //TODO only one row
                sheet_string = "<sheetData>"
             
                    var rowElement = "<row r=" + "\"" + String(row_start) + "\" " + sheet_spans + ">"
                    //Writing start
                    for l in 0..<sheet_location.count {
                        //1 == A1
                        if row_start == Int.parse(from: sheet_location[l])! {
                            if Double(sheet_content[l]) == nil {
                                let sharedLocation = uniqueContent.firstIndex(of: sheet_content[l])
                                if (sharedLocation != nil){
                                    rowElement.append("<c r=" + "\"" + sheet_location[l] + "\" " + "t=\"s\"><v>" + String(sharedLocation!) + "</v></c>")
                                }else{
                                    //TODO? shoud I do something?
                                }
                            }else{
                                rowElement.append("<c r=" + "\"" + sheet_location[l] + "\"><v>" + sheet_content[l] + "</v></c>")
                            }
                        }
                    }
                    
                    rowElement.append("</row>")
                    sheet_string.append(rowElement)
             
                
                sheet_string.append("</sheetData>")
                sheet_data_ary.append(sheet_string)
            }else if row_diff > 0{
                sheet_string = "<sheetData>"
                for k in row_start..<row_end+1 {
                    var rowElement = "<row r=" + "\"" + String(k) + "\" " + sheet_spans + ">"
                    //Writing start
                    for l in 0..<sheet_location.count {
                        //1 == A1
                        if k == Int.parse(from: sheet_location[l])! {
                            if Double(sheet_content[l]) == nil {
                                let sharedLocation = uniqueContent.firstIndex(of: sheet_content[l])
                                if (sharedLocation != nil){
                                    rowElement.append("<c r=" + "\"" + sheet_location[l] + "\" " + "t=\"s\"><v>" + String(sharedLocation!) + "</v></c>")
                                }else{
                                   
                                }
                            }else{
                                rowElement.append("<c r=" + "\"" + sheet_location[l] + "\"><v>" + sheet_content[l] + "</v></c>")
                            }
                        }
                    }
                    
                    rowElement.append("</row>")
                    sheet_string.append(rowElement)
                }
                
                sheet_string.append("</sheetData>")
                sheet_data_ary.append(sheet_string)
            }
            
            
            }
            
            
            
            
        }
        return (sheet_data_ary,sharedString_data)
    }
    
    //TODO
    //"<sst xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" count=\"3\" uniqueCount=\"3\"><si><t>sheet1 empty</t></si><si><t>sheet2 empty</t></si><si><t>sheet3 empty</t></si></sst>"
    
}

extension String {
    var letters: String {
        return String(unicodeScalars.filter(CharacterSet.letters.contains))
    }
}

extension Int {
    static func parse(from string: String) -> Int? {
        return Int(string.components(separatedBy: CharacterSet.decimalDigits.inverted).joined())
    }
}



//defaultRowHeight="13.5"/><sheetData><row r="1" spans="1:1"><c r="A1" t="s"><v>1</v></c></row></sheetData>

//<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
//<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="3" uniqueCount="3">
//    <si>
//        <t>sheet1 empty</t>
//    </si>
//    <si>
//        <t>sheet2 empty</t>
//    </si>
//    <si>
//        <t>sheet3 empty</t>
//    </si>
//</sst>
