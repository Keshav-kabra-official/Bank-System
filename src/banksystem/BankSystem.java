import java.util.Scanner;
///
import java.io.File;
import java.io.IOException;
import jxl.Cell;
import jxl.CellType;
import jxl.Sheet;
import jxl.Workbook;
import jxl.read.biff.BiffException;
import jxl.write.*;
import jxl.write.Number;
///
import java.util.regex.Matcher; 
import java.util.regex.Pattern; 
///
import java.text.SimpleDateFormat;
import java.util.Date;
import java.text.ParseException;


// Shows various bank-names available
// returns the bank_index_number pressed by user
class BankNames
{
    public int show_bank_names()
    {
        Scanner s = new Scanner(System.in);
        System.out.println("Portal to Deposite/Withdraw money from following Banks : ");
        System.out.println("  1. HDFC Bank\n  2. State Bank of India\n  3. Punjab National Bank\n  4. Axis Bank\n  5. Bank of Baroda");
        System.out.println("  6. Indian Bank\n  7. EXIT from Program\n");
        System.out.print("Press shown index numbers to start transactions : ");
        int ch1 = s.nextInt();
        return ch1;
    }
}

// switch-case ladder for chosen banks
// set String-variable to bank name for Opening specified XLS file
class ChosenBank
{
    public String bank_chosen(int ch)
    {
        String bank_name = "";
        switch(ch)
        {
            case 1:
                System.out.println("\nWelcome to HDFC Bank !!!");
                bank_name = "HDFC";
                break;
            case 2:
                System.out.println("\nWelcome to State Bank of India(SBI) !!!");
                bank_name = "SBI";
                break;
            case 3:
                System.out.println("\nWelcome to Panjab Natonal Bank(PNB) !!!");
                bank_name = "PNB";
                break;
            case 4:
                System.out.println("\nWelcome to Axis Bank !!!");
                bank_name = "AxisB";
                break;
            case 5:
                System.out.println("\nWelcome to Bank of Baroda(BOB) !!!");
                bank_name = "BOB";
                break;
            case 6:
                System.out.println("\nWelcome to Indian Bank !!!");
                bank_name = "IndianB";
                break;
            case 7:
                System.out.println("\n --- Program Designed by ---\n\tKESHAV KABRA\n\t+91-7014722936\n\t(keshavkabra118@gmail.com)\n");
                System.exit(0);
            default:
                System.out.println("\nInvalid Key Pressed... Please Press a Valid Key...\n");
                System.exit(0);
        }
        return bank_name;
    }
}

// asking user about her/his Account-Number and Name
class AccountNo_Name_OfUser
{
    public String[] ask_accno_name()
    {
        Scanner s = new Scanner(System.in);
        String[] arr = new String[2];
        System.out.print("Enter Your Account Number : ");
        arr[0] = s.next();
        s.nextLine();  // to take string input after long -- clearing buffer...
        System.out.print("Enter Your Full Name : ");
        arr[1] = s.nextLine();
        return arr;
    }
}


// shows user-account-info as per given account-number and user-name
class ShowAccountInfo
{
    public boolean show_account_info(Workbook wb, Sheet sh, int row, int col, String acc_no, String name)
    {
        boolean flag1 = false;
        String xls_accno1="", xls_name1="";
        for(int i=1;i<row;i++)
        {
            for(int j=0;j<col;j++)
            {
                j++;
                Cell c = sh.getCell(j, i);
                if(j == 1)  xls_accno1 = c.getContents(); // getting account numbers from file
                j++;
                c = sh.getCell(j, i);
                if(j == 2)  xls_name1 = c.getContents(); // getting names from file
                ///
                if(xls_accno1.equals(acc_no) && xls_name1.equals(name))
                {
                    System.out.println("\n--- Hello " + name + ", your account information is : ---");
                    j = 0;
                    int k = 0;
                    flag1 = true;
                    while(j<col)
                    {
                        Cell info = sh.getCell(j, k);
                        System.out.print("  " + info.getContents() + ": ");
                        c = sh.getCell(j,i);
                        System.out.println(c.getContents());
                        j++;
                    }
                    wb.close();
                    break;
                }
                break;
            }
            if(flag1 == true)
                break;
        }
        return flag1;
    }
}


// to deposite money and update XLS file as per given account-number and user-name
class DepositeMoney
{
    public boolean deposite_money_in_account(Workbook wb, Sheet sh, int row, int col, 
            String acc_no, String name, String bank_name, int amt) throws IOException, WriteException
    {
        boolean flag = false;
        String xls_accno="", xls_name="";
        int xls_amount = 0;
        for(int i=1;i<row;i++)
        {
            for(int j=0;j<col;j++)
            {
                j++;
                Cell c = sh.getCell(j, i);
                if(j == 1)  xls_accno = c.getContents(); // getting account numbers from file
                j++;
                c = sh.getCell(j, i);
                if(j == 2)  xls_name = c.getContents(); // getting names from file
                j++;
                c = sh.getCell(j, i);
                if(j == 3)  xls_amount = Integer.parseInt(c.getContents()); // getting current balance from file

                if(xls_accno.equals(acc_no) && xls_name.equals(name))
                {
                    xls_amount += amt;
                    // writing in same xls file and in same sheet
                    WritableWorkbook workbook = Workbook.createWorkbook(new File("C:\\Users\\Keshav\\Desktop\\BankSystem\\"+bank_name+".xls"), wb);
                    WritableSheet sheet = workbook.getSheet(0);

                    WritableCell cell = sheet.getWritableCell(j, i);
                    if(cell.getType() == CellType.NUMBER)
                    {
                        Number n = (Number) cell;
                        n.setValue((double)xls_amount);
                    }
                    
                    System.out.println("\nDear " + xls_name + ", Your account has been Credited with INR " + amt + "/-");
                    System.out.println("--- Money Depostited Successfully ---");
                    workbook.write(); 
                    //Close and free allocated memory 
                    workbook.close(); 
                    wb.close();
                    flag = true;
                }
                break;
            }
            if(flag == true)
                break;
        }
        return flag;
    }
}

// to withdraw money and update XLS file as per given account-number and user-name
class WithdrawMoney
{
    public boolean withdraw_money_from_account(Workbook wb, Sheet sh, int row, int col, 
            String acc_no, String name, String bank_name) throws IOException, WriteException
    {
        Scanner s = new Scanner(System.in);
        boolean flag = false;
        String xls_accno="", xls_name="";
        int xls_amount = 0;
        for(int i=1;i<row;i++)
        {
            for(int j=0;j<col;j++)
            {
                j++;
                Cell c = sh.getCell(j, i);
                if(j == 1)  xls_accno = c.getContents(); // getting account numbers from file
                j++;
                c = sh.getCell(j, i);
                if(j == 2)  xls_name = c.getContents(); // getting names from file
                j++;
                c = sh.getCell(j, i);
                if(j == 3)  xls_amount = Integer.parseInt(c.getContents()); // getting current balance from file


                if(xls_accno.equals(acc_no) && xls_name.equals(name))
                {
                    System.out.print("Enter the Amount(in Indian Rupees) you want to Withdraw : ");
                    int amt_withdraw = s.nextInt();
                    if(xls_amount<=1000 || xls_amount-amt_withdraw<=1000)
                    {
                        System.out.println("\nInsufficient Balance in Account... Can't proceed the transaction ...\n");
                        flag = true;
                        break;
                    }

                    xls_amount -= amt_withdraw;
                    // writing in same xls file and in same sheet
                    WritableWorkbook workbook = Workbook.createWorkbook(new File("C:\\Users\\Keshav\\Desktop\\BankSystem\\"+bank_name+".xls"), wb);
                    WritableSheet sheet = workbook.getSheet(0);

                    WritableCell cell = sheet.getWritableCell(j, i);
                    if(cell.getType() == CellType.NUMBER)
                    {
                        Number n = (Number) cell;
                        n.setValue((double)xls_amount);
                    }

                    System.out.println("\nDear " + xls_name + ", You have Withdrawn INR " + amt_withdraw + "/- from your account...");
                    System.out.println("--- Money Withdrawn Successfully ---");
                    workbook.write(); 
                    //Close and free allocated memory 
                    workbook.close(); 
                    wb.close();
                    flag = true;
                }
                break;
            }
            if(flag == true)
                break;
        }
        return flag;
    }
}

// various account operatons -- account-info, deposite-money, withdraw-money, delete-account, open-account, etc.
// handles entire XLS file
class AccountOperations
{
    //wb, sh, row, col, acc_no, name
    public void account_operations(int choice, Workbook wb, Sheet sh, int row, 
            int col, String acc_no, String name, String bank_name) throws IOException, WriteException
    {
        switch(choice)
        {
            // open a new account
            case 1:
                Open_account obj = new Open_account();
                obj.account_open(bank_name, wb, sh, row, col);
                break;
            
            // shows the details of user
            case 2:
                boolean flag1 = false;
                ShowAccountInfo acc_info = new ShowAccountInfo();
                flag1 = acc_info.show_account_info(wb, sh, row, col, acc_no, name);
                if(flag1 == false)
                {
                    ask_open_account(bank_name, wb, sh, row, col);
                }
                break;
            
            // deposite money in the account 
            case 3:
                Scanner s = new Scanner(System.in);
                boolean flag = false;
                System.out.print("Enter the Amount(in Indian Rupees) you want to Deposite : ");
                int amt = s.nextInt();
                if(amt > 200000)
                {
                    System.out.println("You can not deposite more than (INR)1,99,999/- at one time...");
                    flag = true;
                }
                DepositeMoney money = new DepositeMoney();
                flag = money.deposite_money_in_account(wb, sh, row, col, acc_no, name, bank_name, amt);
                if(flag == false)
                {
                    ask_open_account(bank_name, wb, sh, row, col);
                }
                break;
                
            // withdraw money from the account
            case 4:
                flag = false;
                WithdrawMoney m = new WithdrawMoney();
                flag = m.withdraw_money_from_account(wb, sh, row, col, acc_no, name, bank_name);
                if(flag == false)
                {
                    ask_open_account(bank_name, wb, sh, row, col);
                }
                break;
                
            // delete the account of the user
            case 5:
                flag = false;
                Delete_account o = new Delete_account();
                flag = o.close_account(bank_name, wb, sh, row, col, acc_no, name);
                if(flag == false)
                {
                    ask_open_account(bank_name, wb, sh, row, col);
                }
                break;
            default:
                System.out.println("Invalid Key Pressed... Please Press a Valid Key...\n");
        }
    }
    
    public static void ask_open_account(String bank_name, Workbook wb, Sheet sh, int row, int col) throws IOException, WriteException
    {
        Scanner s = new Scanner(System.in);
        System.out.println("\nYou do not have an Account in " + bank_name + " ...");
        System.out.println("Press 1 to Open an Account ...");
        int ch = s.nextInt();
        if(ch == 1)
        {
            Open_account obj = new Open_account();
            obj.account_open(bank_name, wb, sh, row, col);
        }
    }
}


// open an account
class Open_account
{
    public void account_open(String bank_name, Workbook wb, Sheet sh, int row, int col) throws IOException, WriteException
    {
        Scanner s = new Scanner(System.in);
        /*
            Idea is to take the Id, Account_number, IFSC_number of last entry from XLS file and increment the numeric part of each one,
            to make respective columns for the new entry.
        */
        // Taking data from last entry from XLS file
        int r = row-1;
        Cell ch = sh.getCell(0, r);
        String prev_id = ch.getContents();
        ch = sh.getCell(1, r);
        String prev_accno = ch.getContents();
        ch = sh.getCell(7, r);
        String prev_ifsc = ch.getContents();

        // incrementing numeric part of each fetched data
        long id = Long.parseLong(prev_id);
        id++;
        prev_id = Long.toString(id);
        long acc = Long.parseLong(prev_accno);
        acc++;
        prev_accno = Long.toString(acc);
        String vals[] = prev_ifsc.split("[A-Za-z]+");
        int value = Integer.parseInt(vals[1]);
        value++;
        prev_ifsc = prev_ifsc.substring(0, prev_ifsc.length()-vals[1].length())+ "0" +value;

        s.nextLine();
        ///
        System.out.print("Enter your Full Name : ");
        String new_name = s.nextLine();
        while(new_name.length() < 6)
        {
            System.out.print("Name must be 6 characters long... Please enter a valid name : ");
            new_name = s.nextLine();
        }
        ///
        System.out.print("Enter your Mobile Number : ");
        String new_mobile = s.nextLine();
        while(isValidMobile(new_mobile) == false)
        {
            System.out.print("Invalid Mobile Number... Continue with a valid mobile number : ");
            new_mobile = s.nextLine();
        }
        ///
        System.out.print("Enter your Email Address : ");
        String new_email = s.nextLine();
        while(isValidEmail(new_email) == false)
        {
            System.out.print("Invalid Email... Continue with a valid email address : ");
            new_email = s.nextLine();
        }
        ///
        System.out.print("Enter your Full Resident Address : ");
        String new_addr = s.nextLine();
        ///
        System.out.print("Enter your Sex (F:Female, M:Male, O:Other) : ");
        char new_sex = s.next().charAt(0);
        while(new_sex!='F' && new_sex!='M' && new_sex!='O')
        {
            System.out.print("Enter a Valid Sex Group : ");
            new_sex = s.next().charAt(0);
        }
        ///
        s.nextLine();
        System.out.print("Enter your DOB (DD/MM/YYYY) : ");
        String new_dob = s.nextLine();
        while(isValidDate(new_dob) == false)
        {
            System.out.print("Invalid DOB... Please emter a valid DOB : ");
            new_dob = s.nextLine();
        }
        ///
        System.out.print("Enter the Amount you want to deposite (minimum 1000 INR) : ");
        int bal = s.nextInt();
        while(bal < 1000)
        {
            System.out.print("Minimum required balance is 1000 INR. Please enter a valid amount : ");
            bal = s.nextInt();
        }
        String balance = Integer.toString(bal);

        String[] data = new String[]{prev_id, prev_accno, new_name, balance, new_mobile, new_email, new_addr, prev_ifsc, Character.toString(new_sex), new_dob};

        WritableWorkbook workbook = Workbook.createWorkbook(new File("C:\\Users\\Keshav\\Desktop\\BankSystem\\"+bank_name+".xls"), wb);
        WritableSheet sheet = workbook.getSheet(0);

        r = row;
        for(int j=0;j<col;j++)
        {
            if(j==0 || j==1 || j==3 || j==4)
                sheet.addCell(new Number(j, r, (double)Long.parseLong(data[j])));
            else
                sheet.addCell(new Label(j, r, data[j]));
        }
        
        System.out.println("\nDear " + new_name + ", Your Account has been Openend !!!");
        System.out.println("--- Accound Opened Successfully ---");
        workbook.write();
        //Close and free allocated memory 
        workbook.close(); 
        wb.close();
    }
    
    // function to check validity of email entered
    public static boolean isValidEmail(String email)
    {
        String emailRegex = "^[a-zA-Z0-9_+&*-]+(?:\\."+
                            "[a-zA-Z0-9_+&*-]+)*@" +
                            "(?:[a-zA-Z0-9-]+\\.)+[a-z" +
                            "A-Z]{2,7}$";
        Pattern pat = Pattern.compile(emailRegex);
        if (email == null)
            return false;
        return pat.matcher(email).matches();
    }
    
    //function to check validity of mobile number
    public static boolean isValidMobile(String s)
    {
        // 1) Begins with 0 or +91
        // 2) Then contains 6 or 7 or 8 or 9.
        // 3) Then contains 9 digits
        Pattern p = Pattern.compile("(0/+91)?[6-9][0-9]{9}");
        Matcher m = p.matcher(s);
        return (m.find() && m.group().equals(s));
    }
    
    //function to check validity of DOB
    public static boolean isValidDate(String d)
    {
//        String regex = "^(3[01]"+"|[12][0-9]|0[1-9])/(1[0-2]|0[1-9])/[0-9]{4}$";
//        Pattern pattern = Pattern.compile(regex);
//        Matcher matcher = pattern.matcher((CharSequence)d);
//        return matcher.matches();
        
        /* Check if date is 'null' */
	if (d.trim().equals(""))
	{
	    return false;
	}
	/* Date is not 'null' */
	else
	{
	    SimpleDateFormat sdfrmt = new SimpleDateFormat("dd/MM/yyyy");
	    sdfrmt.setLenient(false);
	    try
	    {
	        Date javaDate = sdfrmt.parse(d); 
	        return true;
	    }
	    /* Date format is invalid */
	    catch(ParseException e)
	    {
	        return false;
	    }
	}
    }
}

// closing an account
class Delete_account
{
    public boolean close_account(String bank_name, Workbook wb, Sheet sh, int row, int col, String acc_no, String name) throws IOException, WriteException
    {
        boolean flag = false;
        String xls_accno="", xls_name="";
        int xls_amount = 0;
        for(int i=1;i<row;i++)
        {
            for(int j=0;j<col;j++)
            {
                j++;
                Cell c = sh.getCell(j, i);
                if(j == 1)  xls_accno = c.getContents(); // getting account numbers from file
                j++;
                c = sh.getCell(j, i);
                if(j == 2)  xls_name = c.getContents(); // getting names from file
                j++;
                c = sh.getCell(j, i);
                if(j == 3)  xls_amount = Integer.parseInt(c.getContents()); // getting current balance from file

                if(xls_accno.equals(acc_no) && xls_name.equals(name))
                {
                    // writing in same xls file and in same sheet
                    WritableWorkbook workbook = Workbook.createWorkbook(new File("C:\\Users\\Keshav\\Desktop\\BankSystem\\"+bank_name+".xls"), wb);
                    WritableSheet sheet = workbook.getSheet(0);

                    sheet.removeRow(i);
                    workbook.write(); 
                    //Close and free allocated memory 
                    workbook.close(); 
                    wb.close();
                    System.out.println("\nDear" + xls_name + ", You will be credited with INR " + xls_amount + "/-");
                    System.out.println("--- Account Closed Successfully ---");
                    flag = true;
                }
                break;
            }
            if(flag == true)
                break;
        }
        return flag;
    }
}


public class BankSystem
{
    public static void main(String[] args) throws Exception, BiffException
    {
        Scanner s = new Scanner(System.in);
        
        // Shows various bank-names available
        // returns the bank_index_number pressed by user
        BankNames obj = new BankNames();
        int ch1 = obj.show_bank_names();
        
        // switch-case ladder for chosen banks
        // set String-variable to bank name for Opening specified XLS file
        ChosenBank ob = new ChosenBank();
        String bank_name = ob.bank_chosen(ch1);
        
        System.out.println("\n  Press :\n    (1) Open Account\n    (2) Account Details\n    (3) Deposite Money\n"
                + "    (4) Withdraw Money\n    (5) Delete Account");
        int ch2 = s.nextInt();
        
        // asking user about her/his Account-Number and Name
        String acc_no="", name="";
        if(ch2 != 1)
        {
            AccountNo_Name_OfUser o = new AccountNo_Name_OfUser();
            String[] st = o.ask_accno_name();
            acc_no = st[0];
            name = st[1];
        }
        
        // READING XLS FILE AND WORKING ON IT ...
        File f = new File("C:\\Users\\Keshav\\Desktop\\BankSystem\\"+bank_name+".xls");
        Workbook wb = Workbook.getWorkbook(f);
        Sheet sh = wb.getSheet(0);
        int row = sh.getRows();
        int col = sh.getColumns();
        
//        PRINTING XLS FILE CONTENTS
//        for(int i=0;i<row;i++)
//        {
//            for(int j=0;j<col;j++)
//            {
//                Cell c = sh.getCell(j,i);
//                System.out.printf("%30s\t", c.getContents());
//            }
//            System.out.println();
//        }
                
        // various account operatons -- account-info, deposite-money, withdraw-money, delete-account, open-account, etc.
        AccountOperations opr = new AccountOperations();
        opr.account_operations(ch2, wb, sh, row, col, acc_no, name, bank_name);
    }
}
