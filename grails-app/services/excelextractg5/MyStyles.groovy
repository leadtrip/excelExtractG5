package excelextractg5

import builders.dsl.spreadsheet.api.Color
import builders.dsl.spreadsheet.builder.api.CanDefineStyle
import builders.dsl.spreadsheet.builder.api.Stylesheet


class MyStyles implements Stylesheet{

    static final HEADERS = 'headers'
    static final SUBHEADING = 'subheading'

    @Override
    void declareStyles(CanDefineStyle stylable) {
        stylable.style(HEADERS) {
            border(bottom) {
                style thick
                color black
            }
            font {
                style bold
            }
            background Color.aliceBlue
        }

        stylable.style(SUBHEADING) {
            font {
                color Color.white
                style bold
            }
            background Color.darkBlue
        }
    }
}
