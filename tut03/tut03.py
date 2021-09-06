 #  KUNAMALLA JOEL RATHNAM 1901EE32

def output_by_subject() :
    import os
    #os.mkdir("output_by_subject")
    try : 
        os.mkdir("output_by_subject")
    except  FileExistsError:
        pass    
    with open("regtable_old.csv", "r") as f:
        for line in f:
            file1 = line.split(',')
            del file1[4:8]
            del file1[2]
            if (file1[2] =="subno"):continue
            try : 
                with open(f"output_by_subject\\{file1[2]}.csv" ):
                    with open(f"output_by_subject\\{file1[2]}.csv", "a") as f11:
                        file1=",".join(file1)
                        f11.write(file1)

            except IOError:
                with open(f"output_by_subject\\{file1[2]}.csv", "w") as f11:
                        f11.write("rollno,register_sem,subno,sub_type\n")
                        file1=",".join(file1)
                        f11.write(file1)
        return


def output_individual_roll():
    import os
    try : 
        os.mkdir("output_individual_roll")
    except  FileExistsError:
        pass
    with open("regtable_old.csv", "r") as f:
        for line in f:
            file1 = line.split(',')
            del file1[4:8]
            del file1[2]
            if (file1[0] =="rollno"):continue
            try: 
                with open(f"output_individual_roll\\{file1[0]}.csv"):
                    with open(f"output_individual_roll\\{file1[0]}.csv", "a") as f11:
                        file1=",".join(file1)
                        f11.write(file1)
            except IOError:
                with open(f"output_individual_roll\\{file1[0]}.csv", "w") as f11:
                        f11.write("rollno,register_sem,subno,sub_type\n")
                        file1=",".join(file1)
                        f11.write(file1)
        return


output_by_subject()
output_individual_roll()