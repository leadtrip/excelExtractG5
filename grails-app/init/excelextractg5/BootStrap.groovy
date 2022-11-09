package excelextractg5

import grails.gorm.transactions.Transactional

class BootStrap {

    def init = { servletContext ->
        setupDb()
    }

    @Transactional
    def setupDb() {
        ('a'..'z').each {
            new Abcd(reqId: 1, question: "qAbcd-$it", value: "aAbcd-$it").save()
            new Wxyz(reqId: 1, question: "qWxyz-$it", value: "aWxyz-$it").save()
        }
    }

    def destroy = {
    }
}
