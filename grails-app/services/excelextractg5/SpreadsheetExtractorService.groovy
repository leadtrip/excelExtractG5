package excelextractg5


import builders.dsl.spreadsheet.builder.poi.PoiSpreadsheetBuilder
import grails.gorm.transactions.Transactional

import javax.servlet.http.HttpServletResponse

@Transactional
class SpreadsheetExtractorService {

    def generate(HttpServletResponse response) {
        log.info("Generating spreadsheet")

        PoiSpreadsheetBuilder.create(response.outputStream).build {
            style ('headers') {
                border(bottom) {
                    style thick
                    color black
                }
                font {
                    style bold
                }
                background whiteSmoke
            }

            sheet('Abcd') {
                // header
                row {
                    cell 'Question'
                    cell 'Value'
                }
                fetchAbcds(1).each {anAbcd ->
                    row {
                        cell anAbcd.question
                        cell anAbcd.value
                    }
                }
            }
            sheet('Wxyz') {
                // header
                row {
                    cell 'Question'
                    cell 'Value'
                }
                fetchWxyzs(1).each {anWxyz ->
                    row {
                        cell anWxyz.question
                        cell anWxyz.value
                    }
                }
            }

            sheet('Random') {
                filter auto             // automatic filters
                row{
                    cell 'Question'
                    cell 'Value'
                }
                row{
                    cell 'Why'
                    cell 'Because'
                }
                row{
                    cell 'Who'
                    cell 'Them'
                }

                row 5, { cell 'Line 5' }    // specify the row number
                row()                                                           // leave row 6 empty
                row { cell 'Line 7' }                       // start again on line 7

                row{
                    cell(2) {                               // specify the cell number, starting at 1
                        value 'Line 8 column (2 or B)'                   // use cell closure, call value
                        width 100                                           // width
                    }
                    cell('C') {                             // specify the cell letter, starting at A
                        formula('YEAR(TODAY())')                            // use cell formula
                    }
                }

                row {                                               // colspan and rowspan
                    cell {
                        value "Columns"
                        colspan 2
                    }
                }
                row {
                    cell {
                        value 'Rows'
                        rowspan 3
                    }
                    cell 'Value 1'
                }
                row {
                    cell ('B') { value 'Value 2' }
                }
                row {
                    cell ('B') { value 'Value 3' }
                }

                row {
                    cell('D') {
                        png image from 'grails-app/assets/images/apple-touch-icon.png'          // add an image
                    }
                }

                // naming cells allows you to refer to them later
                def aRow = row {
                    cell 'A'
                    cell 'B'
                    cell 'A + B'
                }

                row {
                    cell {
                        value 10
                        name '_CellA'
                    }
                    cell {
                        value 20
                        name '_CellB'
                    }
                    cell {
                        formula 'SUM(#{_CellA},#{_CellB})'
                    }
                }
            }
            sheet('Links') {
                row {
                    cell {
                        value 'Go to Random sheet'
                        link to name '_CellA'
                        width auto
                    }
                    cell {
                        value 'File'
                        link to file 'text.txt'
                    }
                    cell {
                        value 'URL'
                        link to url 'https://www.bbc.com'
                    }
                    cell {
                        value 'Mail (plain)'
                        link to email 'someone@example.com'
                    }
                    cell {
                        value 'Mail (with subject)'
                        link to email 'someone@example.com',
                                cc: 'someoneelse@example.com'
                        subject: 'Testing Excel Builder'
                        body: 'It is really great tools'
                    }
                }
            }
        }
    }

    List<Abcd> fetchAbcds( reqId ) {
        Abcd.findAllByReqId( reqId )
    }

    List<Wxyz> fetchWxyzs( reqId ) {
        Wxyz.findAllByReqId( reqId )
    }

}
