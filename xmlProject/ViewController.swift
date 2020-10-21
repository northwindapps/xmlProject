//
//  ViewController.swift
//  xmlProject
//
//  Created by yujin on 2020/10/13.
//  Copyright © 2020 yujin. All rights reserved.
//

import UIKit


class ViewController: UIViewController {
    
    var sheetNumber = 10 // 3 > sheets
    let stringContents = ["tesla","bmw","toyota","toyota","honda","orange","10","renault","nissan","volkswagen","tesla","bmw","1111","$999"]
    
    let locations = ["A4","A1","B2","F4","A4","A1","B2","F4","A4","A1","B3","D5","E7","F2"]
    
    let sheetIdx = [1,2,1,6,5,7,10,4,5,3,4,6,6,10]
    
    let customFileName = "worldCar.xlsx"
    

    override func viewDidLoad() {
        super.viewDidLoad()
        // Do any additional setup after loading the view.
        let serviceInstance = Service(imp_sheetNumber: sheetNumber, imp_stringContents: stringContents, imp_locations: locations, imp_idx: sheetIdx, imp_fileName: customFileName)
        serviceInstance.export()
    
    }
    
    

}
