# program untuk upload data guru
# secara banyak
import openpyxl
import hashlib
import random
import mysql.connector 

listKelasXI = [
    "kelas",
    6,
    7,
    8,
    9,
    1,
    11,
    10,
    2
]


# generate unique code
def generate_unique_code(nis, nama):
    hashes = ""

    randomnumber = random.randrange(0, 100)
    hash_object = hashlib.sha1((nis + str(randomnumber) ).encode('utf-8') )
    pbHash = hash_object.hexdigest()
    hash1 = pbHash
    hash1 = hash1[:24]


    # md5 
    hash_object = hashlib.md5((nis + nama).encode('utf-8'))
    hash2 = hash_object.hexdigest()

    # final sampling
    hash_object = hashlib.md5((hash2 + nama).encode('utf-8'))
    datatoencode = hash_object.hexdigest() + hash1
    hashes = datatoencode

    return hashes



# parsing data 
def parsingDataSource(filename, listKelas, startpoint=1):
    print("[INFO] - Starting reader")

    # Define variable to load the dataframe
    dataframe = openpyxl.load_workbook(filename)

    # Define variable to read sheet
    dataframe1 = dataframe.active

    # temporary data
    listdata = []


    # Iterate the loop to read the cell values
    id_siswa = startpoint - 1
    kelasindex = 0
    for row in range(1, dataframe1.max_row):
        indexcol = 0
        nama = ""
        nomorhp = ""
        jk = ""
        nis = ""
        kodeunik = ""
        id_kelas = ""
        valid = False


        # iterasi pengambilan data
        for col in dataframe1.iter_cols(0, dataframe1.max_column):
            data = (col[row].value)
            

            if data is not None:
                if indexcol == 0 and type(data) == int:
                    # nomor
                    valid = True
                    id_siswa = id_siswa + 1
                    if data == 1:
                        # kelas baru
                        kelasindex = kelasindex + 1


                if valid:
                    if indexcol == 1:
                        # nis
                        data = data.replace(" ", "")
                        dataar = data.split("/")
                        nis = dataar[1]
                    elif indexcol == 2:
                        # nama
                        nama = data
                    elif indexcol == 3 or indexcol == 4:
                        if str(data).upper() == "L":
                            jk = "Laki-laki"
                        else:
                            jk = "Perempuan"



            indexcol = indexcol + 1



        if nama != "" and nis != "" and jk != "":
            nama = nama.upper()
            id_kelas = listKelas[kelasindex]
            kodeunik = generate_unique_code(nis, nama)
            tupdata = (id_siswa, nis, nama, id_kelas, jk, nomorhp, kodeunik )
            listdata.append(tupdata)

            print(tupdata)


    return listdata




# input data
def insertdata(mydb, listdata):
    mycursor = mydb.cursor()

    sql = "INSERT INTO tb_siswa (id_siswa, nis, nama_siswa, id_kelas, jenis_kelamin, no_hp, unique_code) VALUES (%s, %s, %s, %s, %s, %s, %s)"
    val = listdata

    mycursor.executemany(sql, val)

    mydb.commit()

    print(mycursor.rowcount, "telah diinput.") 





# connecting to the database
mydb = mysql.connector.connect(
  host="smkncampalagian.sch.id",
  user="root",
  password="anu",
  database="db_absensi"
)

print(mydb)



# read data
listdata = parsingDataSource("XI.xlsx", listKelasXI, 1)

# kirim data ke server
insertdata(mydb, listdata)