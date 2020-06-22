//
//  ViewController.swift
//  ParseCoreXlsx
//
//  Created by Angelique Babin on 22/06/2020.
//  Copyright © 2020 Angelique Babin. All rights reserved.
//

import UIKit
import CoreXLSX

class ViewController: UIViewController {
    
    @IBOutlet weak var titleLabel: UILabel!
    @IBOutlet weak var skillsLabel: UILabel! // B
    @IBOutlet weak var evidenceSkillsLabel: UILabel! // D
    @IBOutlet weak var listSkillsLabel: UILabel! // A
    @IBOutlet weak var degreeTitleLabel: UILabel! // C
    
    @IBOutlet weak var skillsTextView: UITextView! // B
    @IBOutlet weak var evidenceSkillsTextView: UITextView! // D
    @IBOutlet weak var listSkillsTextView: UITextView! // A
    @IBOutlet weak var degreeLabel: UILabel! // C

    override func viewDidLoad() {
        super.viewDidLoad()
        
        guard let path = Bundle.main.path(forResource: "fichier", ofType: "xlsx"), let file = XLSXFile(filepath: path) else { return }

        do {
//            skillsTextView.text = try "Non-empty cells:\n\n" + file.parseWorksheetPaths()
//                .compactMap { try file.parseWorksheet(at: $0) }
//                .flatMap { $0.data?.rows ?? [] }
//                .flatMap { $0.cells }
//                .map { $0.reference.description }
//                .joined(separator: " ")

            for wbk in try file.parseWorkbooks() {
                for (name, path) in try file.parseWorksheetPathsAndNames(workbook: wbk) {
                    if let worksheetName = name {
                        print("This worksheet has a name: \(worksheetName)")
                        titleLabel.text = worksheetName
                    }

                    let worksheet = try file.parseWorksheet(at: path)
                    for row in worksheet.data?.rows ?? [] {
                        for c in row.cells {
                            print("C : \(c)")
//                            guard let cs = c.s else { return }
        
                            // param xlsx
                            let sharedStrings = try file.parseSharedStrings()
                            guard let columnA = ColumnReference("A") else { return }
                            guard let columnB = ColumnReference("B") else { return }
                            guard let columnC = ColumnReference("C") else { return }
                            guard let columnD = ColumnReference("D") else { return }

                            let columnAStrings = worksheet.cells(atColumns: [columnA])
                              .compactMap { $0.stringValue(sharedStrings) }
                            let columnBStrings = worksheet.cells(atColumns: [columnB])
                              .compactMap { $0.stringValue(sharedStrings) }
                            let columnCStrings = worksheet.cells(atColumns: [columnC])
                              .compactMap { $0.stringValue(sharedStrings) }
                            let columnDStrings = worksheet.cells(atColumns: [columnD])
                              .compactMap { $0.stringValue(sharedStrings) }
                            
                            listSkillsLabel.text = columnAStrings.first
                            skillsLabel.text = columnBStrings.first
                            degreeTitleLabel.text = columnCStrings.first
                            evidenceSkillsLabel.text = columnDStrings.first

                            listSkillsTextView.text = columnAStrings[1]
                            skillsTextView.text = columnBStrings[1]
                            degreeLabel.text = columnCStrings[1]
                            evidenceSkillsTextView.text = columnDStrings[1]
                        }
                    }
                }
            }
        } catch {
            skillsTextView.text = error.localizedDescription
        }
    }
    

}

//  usdRateLabel.text = "c.s : \(cs) - Reference Description : \(c.reference.description) \n \n Info A : \(columnAStrings[0]) \n Info B : \(columnBStrings[0]) \n Info C : \(columnCStrings[0]) \n Info D : \(columnDStrings[0]) \n \n A : \(columnAStrings[1]) \n B : \(columnBStrings[1]) \n C : \(columnCStrings[1]) \n D : \(columnDStrings[1]) \n \n A : \(columnAStrings[15]) \n B : \(columnBStrings[15]) \n C : \(columnCStrings[15]) \n D : \(columnDStrings[15])"
