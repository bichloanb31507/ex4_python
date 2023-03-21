import traceback
from dao.StudentInterface import StudentInterface
from db import DbContext
from models.Student import Student
from utils import Constant, FileIO

class StudentDAO(StudentInterface):
    
    def __init__(self):
        self.conn = DbContext.connect();
    
    def addStudent(self, student):
        try:
            conn = self.conn
            cursor = conn.cursor()
            values = (student.id, student.code, student.firstName, student.lastName, student.birthOfDate, student.math, student.physics, student.chemistry)
            cursor.execute(Constant.ADD_SQL ,values)
            conn.commit()
            print("addStudent thanh cong")
        except:
            if conn is not None:
                conn.rollback()
                conn.close()
            traceback.print_exc()
            print("addStudent that bai")
    
    def updateStudent(self, student):
        try:
            conn = self.conn
            cursor = conn.cursor()
            values = (student.firstName, student.lastName, student.birthOfDate, student.math, student.physics, student.chemistry, student.id)
            cursor.execute(Constant.UPDATE_SQL, values)
            conn.commit()
            print("updateStudent thanh cong")
        except:
            if conn is not None:
                conn.rollback()
                conn.close()
            traceback.print_exc()
            print("updateStudent that bai")
    
    def deleteStudent(self, student_id):
        try:
            conn = self.conn
            cursor = conn.cursor()
            values = (student_id,)
            cursor.execute(Constant.DELETE_SQL, values)
            conn.commit()
            print("deleteStudent thanh cong")
        except:
            if conn is not None:
                conn.rollback()
                conn.close()
            traceback.print_exc()
            print("deleteStudent that bai")
    
    def getStudent(self, student_id):
        try:
            conn = self.conn
            cursor = conn.cursor()
            values = (student_id,)
            cursor.execute(Constant.GET_STUDENT_SQL, values)
            result = cursor.fetchone()
            if result:
                print(result)
                return Student(*result)
            else:
                return None
        except:
            if conn is not None:
                conn.rollback()
                conn.close()
            traceback.print_exc()
    
    def getAll(self):
        try:
            conn = self.conn
            cursor = conn.cursor()
            cursor.execute(Constant.GET_ALL_SQL)        
            results = cursor.fetchall()
            students = []
            for result in results:
                print(result)
                students.append(Student(*result));
        except:
            if conn is not None:
                conn.rollback()
                conn.close()
            traceback.print_exc()