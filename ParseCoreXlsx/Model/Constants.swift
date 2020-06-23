//
//  Constants.swift
//  ParseCoreXlsx
//
//  Created by Angelique Babin on 23/06/2020.
//  Copyright © 2020 Angelique Babin. All rights reserved.
//

import Foundation
import CoreXLSX

// MARK: - Trick to get static variable in Swift

struct Constants {
    static let columnA = ColumnReference("A")
    static let columnB = ColumnReference("B")
    static let columnC = ColumnReference("C")
    static let columnD = ColumnReference("D")
}
