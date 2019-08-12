import org.docx4j.openpackaging.packages.WordprocessingMLPackage
import org.docx4j.wml.ContentAccessor
import org.docx4j.wml.Tbl
import org.docx4j.wml.Text
import org.docx4j.wml.Tr
import java.io.File
import javax.xml.bind.JAXBElement
import kotlin.reflect.KClass

class UpdateWordTemplate {

    
    private var result: ArrayList<Any> = ArrayList()

    /**
     * Method loads word docx as a WordprocessingMLPackage for further manipulation
     * @param filePath: location path to word document which will be used
     */
    fun getTemplate(filePath: InputStream): WordprocessingMLPackage {
        return WordprocessingMLPackage.load(filePath)
    }

    /**
     * This operation is a wrapper around a couple of JAXB operations that allows you to search through
     * a specific element and all it's children for a certain class.
     * @param obj: element in which we will search through for a wanted class
     * @param toSearch: class of the element we want to find
     * @return List of elements that match the class
     */
    private fun getAllElementsFromObject(obj: Any, toSearch: KClass<*>): List<Any> {

        var objValue: Any = obj

        if (obj is JAXBElement<*>) {
            objValue = obj.value
        }

        if (objValue.javaClass == toSearch.java) {
            result.add(objValue)
            return result
        } else if (objValue is ContentAccessor) {
            val children: List<Any> = objValue.content
            for (child in children) {
                getAllElementsFromObject(child, toSearch)
            }
        }
        return result
    }

    /**
     * This method takes in the element and populate it with our data
     * @param obj: element in which placeholders will be replaced with our data
     * @param valuesToAdd: Map of text and their corresponding placeholders
     * @param placeholders: Array of placeholders we want to replace with our data
     */
    private fun replacePlaceholder(obj: Any, valuesToAdd: List<Map<String, String>>, placeholders: Array<String>) {
        result = ArrayList()
        var columns: List<Any> = getAllElementsFromObject(obj, Text::class)

        for (col in columns) {
            var textElement: Text = col as Text
            for (placeholder in placeholders) {
                if (textElement.value == placeholder) {
                    for (values in valuesToAdd) {
                        textElement.value = values[placeholder]
                    }
                }
            }
        }
    }

    /**
     * Write the document back to a file
     * @param template: our modified version of word file
     * @param filePath: path to the original word file which will be updated
     */
    fun writeDocx(template: WordprocessingMLPackage, filePath: String) {
        val file = File(filePath)
        template.save(file)
    }

    /**
     * Method finds the table in which we have placeholders we wish to replace, finds all rows in that table and invokes method
     * which replaces those placeholders with our data
     * @param placeholders: array of strings we wish to replace with our data
     * @param replacementText: map of strings we wish to write to a file with their corresponding placeholders
     * @param template: template, as a copy of our word file in which we find table
     */
    private fun replaceTable(
        placeholders: Array<String>,
        replacementText: List<Map<String, String>>,
        template: WordprocessingMLPackage
    ) {
        val tables: List<Any> = getAllElementsFromObject(template.mainDocumentPart.contents.body, Tbl::class)
        var tableToAdd: Tbl = Tbl()
        for (values in replacementText) {

            //1. find table
            var tempTable: Tbl = getTemplateTable(tables as List<Tbl>, placeholders[0])
            if (tempTable.content.isEmpty()) {
                tempTable = createTable(template, tableToAdd)
            } else {
                tableToAdd = tempTable
            }
            result = ArrayList()
            //Find rows in which we will replace strings
            var rows: List<Any> = getAllElementsFromObject(tempTable, Tr::class)

            if (rows.isNotEmpty()) {
                // replace each row placeholder with our text
                for (row in rows) {
                    replacePlaceholder(row, listOf(values), placeholders)
                }
            }
        }

    }

    /**
     * Method checks whether a table contains one of our placeholders. If so that table is returned.
     * @param tables: list of tables found in template (word doc)
     * @param placeholdersInTable: placeholders of the table we want to find
     * @return Tbl
     */
    private fun getTemplateTable(tables: List<Tbl>, placeholdersInTable: String): Tbl {
        var templateTable = Tbl()

        for (table in tables) {
            result = ArrayList()
            var textElements: List<Any> = getAllElementsFromObject(table, Text::class)

            for (text in textElements) {
                var textElement: Text = text as Text
                if (textElement.value != null && textElement.value == placeholdersInTable) {
                    templateTable = table
                    return templateTable
                }

            }
        }
        return templateTable
    }

    private fun createTable(wordPackage: WordprocessingMLPackage, tbl: Tbl): Tbl {
        wordPackage.mainDocumentPart.addObject(tbl)
        return tbl
    }

    fun updateWordDocument(
        placeholders: Array<String>,
        replacementText: List<Map<String, String>>,
        template: WordprocessingMLPackage,
        filePath: String
    ) {
        replaceTable(placeholders, replacementText, template)
        writeDocx(template, filePath)
    }

}
}



