import argparse
import lxml.etree
import sys
import zipfile

class OfficeData:
    def __init__(self, item):
            self.item = item
            self.getMetadata(self.item)
            
    def getMetadata(self, file):
        """
        Parsing core.xml file and app.xml file
        :param: MS Office file path
        """
        docxData = None
        
        try:
            docxData = zipfile.ZipFile(file)

        except zipfile.BadZipfile as zipErr:
            print("{} 는 zip 파일이 아닙니다.".format(file))

            if self.isPKFile(self.item):
                print("{} 는 MS office 파일이 아닙니다.".format(file))

            sys.exit()

        print("\n\n=====================================")
        print("‖          Parsing Results          ‖")
        print("=====================================")

        with docxData.open("docProps/core.xml") as core:
            print("\n----------Basic information----------\n")
            print("File : {}".format(file))
            self.parseXML(core)

        with docxData.open("docProps/app.xml") as app:
            print("-------Application Information-------\n")
            self.parseXML(app)

    def isPKFile(self, file):
        """
        Check if it's a valid MS office file
        :param: file path
        :return: True if it's valid, if not, return False
        """
        with open(file, "rb") as fileMagicNumber:
            if fileMagicNumber.read(2) != b"PK":
                return True

        return False

    def parseXML(self, xmlFile):
        """
        Parses OpenXML compressed file
        :param: xmlFile name
        :return: dictionary containing XML file information
        """

        xmlElementTree = lxml.etree.parse(xmlFile)
        foundData = dict()

        for element in xmlElementTree.iter("*"):
            tagName = lxml.etree.QName(element.tag).localname
            if element.text is not None:
                foundData[tagName] = element.text

        self.printData(foundData)
        print("\n")

    def printData(self, foundData):
        """
        Prints found data
        :param: Gathered elements
        """
        for key, value in foundData.items():
            print("{0}: {1}".format(key.title(), value))
        
def main():
    parser = argparse.ArgumentParser(description="Microsoft docx file metadata extractor.")

    parser.add_argument("-m", "--media", dest="media", action="store_false",
                    help="Do not uncompress the stored media in the specified directory.")


    parser.add_argument("-d", "--directory", dest="directory", action="store", default=".",
                        help="Name of the  directory where to output the media.")

    parser.add_argument(dest="item", action="store", metavar="[file or directory]",
                        help="DOCX file or directory to parse.")

    args = parser.parse_args()
    
    OfficeData(args.item)

if __name__ == '__main__':
    main()
