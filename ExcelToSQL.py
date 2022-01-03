class ExcelToSQL:
    def __init__(self,txtFileName='sql.txt'):
        self.filename="data.xlsx"
        self.txtFileName= txtFileName
        self.tableName=""
        self.operation="i"
        self.refCol=""
        self.updatingCol=[]
        self.noOfRows=0
        self.exportSQLTxtFile="sql.txt"
        if(len(self.filename.strip())== 0):
            self.filename="data.xlsx"
        self.EmptyTheTextFile()
        self.TakeUserRequirements()
        self.CheckWhichOperationToRun()

    def AppendSQLToText(self,txtfilename,contentToWrite):
        file_object = open(txtfilename, 'a')
        file_object.write(contentToWrite)
        file_object.close()

    def EmptyTheTextFile(self):
        file = open(self.txtFileName,"w")
        file.close()

    def TakeUserRequirements(self):
        result=[]
        msg="""
        ===========================================================

        Welcome to Excel data extract to Sql
        Please select a action

        1 Update into Database
        2 Insert data into Database

        Give u for update operation or i for insert operation

        ===========================================================

        """
        print(msg)
        self.operation=input("        Enter your action   ").lower()
        msg="""

        ===========================================================

        Now write the name of the Excel file in .xlsx format
        to read data from.

        example: data.xlsx

        ===========================================================
        """
        print(msg)
        self.filename=input("        Enter the excel filename   ")

        msg="""

        =================================================================================
        Enter the reference table name
        (note: the reference table name should be similar with the table in the database)
        =================================================================================

        """
        print(msg)
        self.tableName=input("        Enter the reference table name  ")





    def CheckWhichOperationToRun(self):
        if(self.operation=="u"):
            self.Update()
        elif(self.operation=="i"):
            self.Insert()


    #------Run operation for Update action
    #------Takes three input from user table name, ref col name,updating col name

    def UpdateOperation(self):
        cols=[]
        msg="""

        ====================================================================================
        Enter the reference column name
        (note: the reference column name should be similar with the column with the database
        Its possible to enter multiple column name separated with comma)
        ====================================================================================

        """
        print(msg)
        x=input("        Enter the reference column name  ")
        self.refCol=x
        cols.append(x)
        msg="""

        Enter the updating column name
        (note: the column that should be updated with data, must have the same name with the database)

        """
        print(msg)
        x=input("        Enter the updating column name  ")
        x=x.split(",")
        self.updatingCol=x
        cols.append(x)
        return cols


    def GetValidRefAndUpdatingColName(self,cols):
        try:
            print('===========================================')
            for y in cols[1]:
                print(y+" exist in excel file.")
            print(self.refCol+" exist in excel file.")
            print('===========================================')

        except KeyError:
            print('No information on column'+y+'found')
            cols=self.UpdateOperation()
            self.GetValidRefAndUpdatingColName(cols)


    # ---------This is the main function to run the update operation and return sql text

    def Update(self):
        self.GetValidRefAndUpdatingColName(cols=self.UpdateOperation())
        data=self.GenerateSQLForUpdate()

    # --------Return data like list of rows to update in db,list of column name, list of reference column value

    def GetDataForUpdate(self):
        import pandas as pd
        df = pd.read_excel(self.filename)
        df = df.fillna('N/A')
        d=df.columns.tolist()
        shape=df.shape
        self.noOfRows=shape[0]
        rowData=[]
        rows=[]
        colPos=[]

        for i in range(len(self.updatingCol)):
            pos=d.index(self.updatingCol[i])
            colPos.append(pos)

        for x in range(shape[0]):
            for y in range(len(colPos)):
                rowData.append(df.loc[x][colPos[y]])
            rows.append(rowData)
            rowData=[]
        #-----all rows data in 2d list--------column name list-------listof referenced column id
        return rows,d,df[self.refCol].tolist()


    #---------Return a single line query.
    #---------Takes tablesname, single valuetoUpdate condition string, respective reference column value

    def SingleLineUpdateQuery(self,tableName,valuesToUpdate,refColValue):
        qry="UPDATE "+tableName.upper()+" SET "+valuesToUpdate+" WHERE "+self.refCol+" = "+str(refColValue)+";"
        return qry

    #---------Renders each single line querys for given data as input parameter
    #---------Takes value which to update array, reference column value array

    def GenerateSQLForUpdate(self):
    #try:
        data=self.GetDataForUpdate()
        rowsOfValuesToUpdate=self.GenerateValuesToUpdate(self.updatingCol,data[0])
        for x in range(len(rowsOfValuesToUpdate)):
            x=self.SingleLineUpdateQuery(self.tableName,rowsOfValuesToUpdate[x],data[2][x])
            x=x+"\n"
            self.AppendSQLToText(self.exportSQLTxtFile,x)
        print("Added Sql query in "+self.txtFileName+" Successfully !")
    #except:
        #print("Failed to Generate SQL.")



    #--------Returns array of values to update so that those might be concatanated with other SQL string
    #--------Takes list of columns names and data to insert
    def GenerateValuesToUpdate(self,columnNames,datas):
        sqlText=""
        rowOfValuesToUpdate=[]
        for j in range(len(datas)):
            for i in range(len(columnNames)):
                if(isinstance(datas[j][i], str)):
                    firstPart=columnNames[i]+"='"+datas[j][i]+"'"
                else:
                    firstPart=columnNames[i]+"="+str(datas[j][i])
                if(columnNames[i]!=columnNames[len(columnNames)-1]):
                    firstPart+=","
                sqlText+=firstPart
            rowOfValuesToUpdate.append(sqlText)
            sqlText=""
        return rowOfValuesToUpdate

    #Insert Section Code

    # Get all the rows from excel
    # Here the first column would be all the columns name
    # It returns 2d array of data extracted
    def GetDataForInsert(self):
        import pandas as pd
        df = pd.read_excel(self.filename)
        df = df.fillna('N/A')
        shape=df.shape
        rowData=[]
        rows=[]
        d=df.columns.tolist()
        for x in d:
            rowData.append(x)
        rows.append(rowData)
        rowData=[]
        for x in range(shape[0]):
            for y in range(shape[1]):
                rowData.append(df.loc[x][y])
            rows.append(rowData)
            rowData=[];
        return rows

    #---------Return a single line query for Insert.
    #---------Takes middleQuery, endquery as parameter

    def SingleLineInsertQuery(self,middleQuery,endQuery):
        qry="INSERT INTO "+self.tableName+"("+middleQuery+")"+"VALUES ("+endQuery+");"
        return qry

    def MiddleQueryAndEndQueryGenerator(self,data):
        #generate middlequery
        middleQuery=""
        endQuery=""
        endQueries=[]

        for x in data[0]:
            middleQuery+=x
            if(x!=data[0][len(data[0])-1]):
                middleQuery+=","
        for y in data[1:]:
            for w in y:
                if(isinstance(w, str)):
                    endQuery+="'"+str(w)+"',"
                else:
                    endQuery+=str(w)+","
            endQueries.append(endQuery[:-1])
            endQuery=""
        return middleQuery,endQueries

    def Insert(self):
        try:
            receivedData=self.MiddleQueryAndEndQueryGenerator(self.GetDataForInsert())
            for x in receivedData[1]:
                sqlQuery=self.SingleLineInsertQuery(receivedData[0],x)
                sqlQuery+="\n"
                self.AppendSQLToText('sql.txt',sqlQuery)
            print("Added Sql query in "+self.txtFileName+" Successfully !")
        except:
            print("Failed to Generate Sql")
