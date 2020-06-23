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
    
    // MARK: - Outlets
    
    @IBOutlet weak var titleLabel: UILabel!
    @IBOutlet weak var skillsLabel: UILabel! // B
    @IBOutlet weak var evidenceSkillsLabel: UILabel! // D
    @IBOutlet weak var listSkillsLabel: UILabel! // A
    @IBOutlet weak var degreeTitleLabel: UILabel! // C
    
    @IBOutlet weak var skillsTextView: UITextView! // B
    @IBOutlet weak var evidenceSkillsTextView: UITextView! // D
    @IBOutlet weak var listSkillsTextView: UITextView! // A
    @IBOutlet weak var degreeLabel: UILabel! // C
    
    // MARK: - Properties
    
    var columns = Columns()

    // MARK: - View Life Cycle

    override func viewDidLoad() {
        super.viewDidLoad()
        
        guard let path = Bundle.main.path(forResource: "fichier", ofType: "xlsx"), let file = XLSXFile(filepath: path) else { return }

        do {
            try parseFile(file)
        } catch {
            skillsTextView.text = error.localizedDescription
        }
    }
    
    // MARK: - Methods
    
    private func parseFile(_ file: XLSXFile) throws {
        for wbk in try file.parseWorkbooks() {
            for (name, path) in try file.parseWorksheetPathsAndNames(workbook: wbk) {
                if let worksheetName = name {
                    print("This worksheet has a name: \(worksheetName)")
                    titleLabel.text = worksheetName
                }
                let worksheet = try file.parseWorksheet(at: path)
                for row in worksheet.data?.rows ?? [] {
                    for cell in row.cells {
                        print(cell)
                        setColumnsFile(file: file, worksheet: worksheet)

                        setLabels(Columns(columnAStrings: columns.columnAStrings,
                                          columnBStrings: columns.columnBStrings,
                                          columnCStrings: columns.columnCStrings,
                                          columnDStrings: columns.columnDStrings))
                        setInfosFile(Columns(columnAStrings: columns.columnAStrings,
                                             columnBStrings: columns.columnBStrings,
                                             columnCStrings: columns.columnCStrings,
                                             columnDStrings: columns.columnDStrings))
                    }
                }
            }
        }
    }
    
    private func setColumnsFile(file: XLSXFile, worksheet: Worksheet) {
        do {
            let sharedStrings = try file.parseSharedStrings()
            guard let columnA = Constants.columnA else { return }
            guard let columnB = Constants.columnB else { return }
            guard let columnC = Constants.columnC else { return }
            guard let columnD = Constants.columnD else { return }

            columns.columnAStrings = worksheet.cells(atColumns: [columnA])
                .compactMap { $0.stringValue(sharedStrings) }
            columns.columnBStrings = worksheet.cells(atColumns: [columnB])
                .compactMap { $0.stringValue(sharedStrings) }
            columns.columnCStrings = worksheet.cells(atColumns: [columnC])
                .compactMap { $0.stringValue(sharedStrings) }
            columns.columnDStrings = worksheet.cells(atColumns: [columnD])
                .compactMap { $0.stringValue(sharedStrings) }
        } catch {
            skillsTextView.text = error.localizedDescription
        }
    }
    
    private func setLabels(_ columns: Columns) {
        listSkillsLabel.text = columns.columnAStrings.first
        skillsLabel.text = columns.columnBStrings.first
        degreeTitleLabel.text = columns.columnCStrings.first
        evidenceSkillsLabel.text = columns.columnDStrings.first
    }
//
    private func setInfosFile(_ columns: Columns) {
        listSkillsTextView.text = columns.columnAStrings[1]
        skillsTextView.text = columns.columnBStrings[1]
        degreeLabel.text = columns.columnCStrings[1]
        evidenceSkillsTextView.text = columns.columnDStrings[1]
    }

//    private func setLabels(_ columnAStrings: [String], _ columnBStrings: [String], _ columnCStrings: [String], _ columnDStrings: [String]) {
//        listSkillsLabel.text = columnAStrings.first
//        skillsLabel.text = columnBStrings.first
//        degreeTitleLabel.text = columnCStrings.first
//        evidenceSkillsLabel.text = columnDStrings.first
//    }

//    private func setInfosFile(_ columnAStrings: [String], _ columnBStrings: [String], _ columnCStrings: [String], _ columnDStrings: [String]) {
//        listSkillsTextView.text = columnAStrings[1]
//        skillsTextView.text = columnBStrings[1]
//        degreeLabel.text = columnCStrings[1]
//        evidenceSkillsTextView.text = columnDStrings[1]
//    }
}

// // // // // // // // // // // // // // // // // // // // // // // // //

//  testlabel.text = "c.s : \(cs) - Reference Description : \(c.reference.description) \n \n Info A : \(columnAStrings[0]) \n Info B : \(columnBStrings[0]) \n Info C : \(columnCStrings[0]) \n Info D : \(columnDStrings[0]) \n \n A : \(columnAStrings[1]) \n B : \(columnBStrings[1]) \n C : \(columnCStrings[1]) \n D : \(columnDStrings[1]) \n \n A : \(columnAStrings[15]) \n B : \(columnBStrings[15]) \n C : \(columnCStrings[15]) \n D : \(columnDStrings[15])"

//                        guard let cellS = cell.s else { return }

//            skillsTextView.text = try "Non-empty cells:\n\n" + file.parseWorksheetPaths()
//                .compactMap { try file.parseWorksheet(at: $0) }
//                .flatMap { $0.data?.rows ?? [] }
//                .flatMap { $0.cells }
//                .map { $0.reference.description }
//                .joined(separator: " ")

//            guard let columnA = ColumnReference("A") else { return }
//            guard let columnB = ColumnReference("B") else { return }
//            guard let columnC = ColumnReference("C") else { return }
//            guard let columnD = ColumnReference("D") else { return }

//                        setLabels(columnAStrings, columnBStrings, columnCStrings, columnDStrings)
//                        setInfosFile(columnAStrings, columnBStrings, columnCStrings, columnDStrings)
