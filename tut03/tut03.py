 #  KUNAMALLA JOEL RATHNAM 1901EE32

def output_by_subject() :
    import os
    #os.mkdir("output_by_subject")
    if(os.path.exists("output_by_subject")) : 
        pass
    else :
            os.mkdir("output_by_subject")
    with open("regtable_old.csv", "r") as f:
        for line in f:
            file1 = line.split(',')
            del file1[4:8]
            del file1[2]
            if (file1[2] =="subno"):continue
            if(os.path.exists(f"output_by_subject\\{file1[2]}.csv")): 
                    with open(f"output_by_subject\\{file1[2]}.csv", "a") as f11:
                        file1=",".join(file1)
                        f11.write(file1)

            else :
                with open(f"output_by_subject\\{file1[2]}.csv", "w") as f11:
                        f11.write("rollno,register_sem,subno,sub_type\n")
                        file1=",".join(file1)
                        f11.write(file1)
        return


def output_individual_roll():
    import os
    if(os.path.exists("output_individual_roll")) : 
        pass
    else  :
        os.mkdir("output_individual_roll")
    with open("regtable_old.csv", "r") as f:
        for line in f:
            file1 = line.split(',')
            del file1[4:8]
            del file1[2]
            if (file1[0] =="rollno"):continue
            if(os.path.exists(f"output_individual_roll\\{file1[0]}.csv")) :
                with open(f"output_individual_roll\\{file1[0]}.csv", "a") as f11:
                        file1=",".join(file1)
                        f11.write(file1)
            else :
                with open(f"output_individual_roll\\{file1[0]}.csv", "w") as f11:
                        f11.write("rollno,register_sem,subno,sub_type\n")
                        file1=",".join(file1)
                        f11.write(file1)
        return


output_by_subject()
output_individual_roll()