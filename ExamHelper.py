import docx
import glob
import os

class search:
    fullText, lists, toRemove = [], [], []
    searchWord = ""

    def __init__(self):      
        self.searchWord = input("Search word/sentence:").lower()

    def info(self,msg):
        print("\n[INFO]:"+msg)

    def do_everything_auto(self):
        self.load_files()
        self.get_lists()
        self.fix_fulltext()

    def load_files(self):
        #It will return all docx files from related directory.
        documents = glob.glob((os.getcwd())+"/*.docx")
        for doc in documents:
            try:
                file = docx.Document(doc)
                for txt in file.paragraphs:
                    self.fullText.append("\t"+(txt.text).lower())
            except:
                self.info("Propably some docx files are empty, not looking in them...")
    
    def get_lists(self):
        for i in range(0,len(self.fullText)):
            toAdd= ""
            if "<list" in self.fullText[i]:
                a=i+1
                while ">" not in self.fullText[a]:
                    toAdd += self.fullText[a]+"\n"
                    self.toRemove.append(a)
                    a += 1

                toAdd += self.fullText[a][0:-1]
                self.toRemove.append(i)
                self.lists.append(toAdd+"\n")
                i = a+1

    def fix_fulltext(self):
        removed = 0
        for num in self.toRemove:
            self.fullText.remove(self.fullText[num-removed])
            removed += 1

    def show_results(self):
        self.paragraph_results()
        self.list_results()

    def paragraph_results(self):
        if len(self.fullText) == 0:
            self.info("No paragraph to search...")
            return
        self.info("Paragraph Results For [{}]:".format(self.searchWord))
        for p in self.fullText:
            if (self.searchWord in p):
                print(p[1:],"\n")

    def list_results(self):
        if len(self.lists) == 0:
            self.info("No list to search...")
            return
        self.info("List Results For [{}]:".format(self.searchWord))
        for l in self.lists:
            if self.searchWord in l:
                print(l+"\n")


searcher = search()
searcher.do_everything_auto()
searcher.show_results()

while True:
    searcher.__init__()
    searcher.show_results()