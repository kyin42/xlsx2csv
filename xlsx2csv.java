
import java.io.*;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.TimeZone;
import java.lang.Character;

// First is the SharedStringFile, Second is the OutTextFile and the others are the Sheets (1 or more)
public class testing {
	public static void main(String[] args) throws IOException {
		FileReader in = null;
		String[] infiles = new String[args.length - 3];
		String styles = args[0];
		String sharedStrings = args[1];
		String outFile = args[2];
		for(int indx=0;indx<infiles.length;indx++){
			infiles[indx] = args[indx + 3]; // Deal with the In and Out files
		}

        System.out.println("Starting Extract Text from xlsx file");

		FileWriter out = null;
		int isString = 0;
		int d = 0;
		int tag = 0;
		if (tag == 0) {
			tag = 0;
		}
		int dataNumber = 0;
		int insheet = 0;
		int inrow = 0;
		int needline = 0;
		int isNum = 0;
		int isData = 0;
		int isDate = 0;
		int maxColumn;
		int lastColumn = 1;
		int isGoodStyleData = 0;
		String[] dataArray;
		String tmp_String = null;
		String data = null;
		String width = null;
		String[] uniqueStrings = null;
		String datestyle = "";
		try {
			in = new FileReader(styles);
			int c;
			while ((c = in.read()) != -1) {
				if (c == Character.valueOf('<')) {
					data = null;
					tag = 1;
				} else if (c == Character.valueOf('>')) {
					if(data.startsWith("cellXfs")) {
						isGoodStyleData = 1;
					}
					if(data.startsWith("/cellXfs")) {
						isGoodStyleData = 0;
					}
					if(data.startsWith("xf")) {
						if (isGoodStyleData == 1){
							dataArray = data.split("\"");
							int dateTag = 0;
							for(int i = 0; i<dataArray.length; i++){
								if(dataArray[i].endsWith("numFmtId=")){
									int value = Integer.parseInt(dataArray[i+1]);
									if ( value == 19 || value == 20 || value == 21 || value == 22){
										dateTag = 1;
									}
									
								}
								if (dateTag == 1){	
									if(dataArray[i].endsWith("applyNumberFormat=")){
										datestyle += Integer.parseInt(dataArray[i+1]);
										datestyle += " ";
									}
								}
							}
						}
					}
					if(data.startsWith("sst")) {
						uniqueStrings = new String[Integer.parseInt(data.split("\"")[(data.split("\"").length - 1)])];
					}
					data = null;
					
				} else {
					if (data == null){
						data = "";
						data += (char)c;
					} else {
						data += (char)c;
					}
				}
			}
			c = 0;
			d = 0;
			dataNumber = 0;
			data = null;
			tag = 0;;
			uniqueStrings = null;
			
			
			
			in = new FileReader(sharedStrings);
			out = new FileWriter(outFile);
			
			
			while ((c = in.read()) != -1) {
				if (c == Character.valueOf('<')) {
					if (d == 1) {
						d = 0;
						uniqueStrings[dataNumber] = data;
						dataNumber += 1;
					}
					data = null;
					tag = 1;
				} else if (c == Character.valueOf('>')) {
					if(data.startsWith("sst")) {
						uniqueStrings = new String[Integer.parseInt(data.split("\"")[(data.split("\"").length - 1)])];
					}
					if(data.startsWith("t")) {
						d = 1;
					}
					data = null;
					
				} else {
					if (data == null){
						data = "";
						data += (char)c;
					} else {
						data += (char)c;
					}
				}
			}
			/*FileWriter file = new FileWriter("uniqueStrings.txt");
			for (int l = 0;l<uniqueStrings.length;l++){
			file.write(uniqueStrings[l]);
			file.write("|");
			
			//file.write(Arrays.toString(uniqueStrings));
			}*/
			//System.out.println(uniqueStrings[1392]);
			for (int i = 0;i<infiles.length;i++){
				System.out.println("Processing : Work Sheet " + (i + 1) );
				in = new FileReader(infiles[i]);
				tag = 0;
				d = 0;
				data = null;
				while ((c = in.read()) != -1) {
					if(c == Character.valueOf('<')) {
						if (isData == 1) {
							isData = 2;
						}
						if (isData == 2 && data != null) {
							if (needline == 1) {
								out.write("|");
							}
							if (isDate == 1) {					
								double dou = Double.valueOf(data);
								//d -= 25569;
								dou -= 25569;
								dou *= 86400000;
								long l = (long)Math.round(dou);
								Date date = new Date(l);
								SimpleDateFormat dateFormat = new SimpleDateFormat("yyyy-MM-dd HH-mm-ss");
								dateFormat.setTimeZone(TimeZone.getTimeZone("UTC"));
								out.write(dateFormat.format(date));
							} else if (isNum == 1) {
								if (data.contains("E") || data.contains(".")) {
                                    tmp_String = String.format("%.5f",Double.valueOf(data));
									out.write(tmp_String);
								} else {
									out.write(data);
								}
							} else if (isString == 1) {
								out.write(uniqueStrings[Integer.parseInt(data)]);
							}
							isString = 0;
							isNum = 0;
							isDate = 0;
							needline = 1;
						}
						tag = 1;
						data = null;
					} else if (c == Character.valueOf('>')) {
						if (data.startsWith("dimension")) {
							width = data.split("\"")[(data.split("\"").length - 2)].split(":")[1].replaceAll("\\d*$", "");
							//System.out.println(Character.getNumericValue('a') - 9);
							//System.out.println(Character.getNumericValue('A') - 9);
							//System.out.println(Character.getNumericValue('z') - 9);
							//System.out.println(Character.getNumericValue('Z') - 9);//34
							int jk = 0;
							for(int jkk = 0;jkk<width.length();jkk++)
							{
								if(jkk == 0 && width.length() == 2){
									jk += (Character.getNumericValue(width.charAt(jkk)) - 9) * 26;
							
								}else{
									jk += Character.getNumericValue(width.charAt(jkk)) - 9;
								}
							}
							maxColumn = jk;
							//System.out.println(jk);
						}
						if (data.startsWith("sheetData")) {
							insheet = 1;
						}
						if (insheet == 1) {
							if (data.startsWith("row")) {
								inrow = 1;
							}
							if (inrow == 1) {
								if (data.startsWith("c")) {
									width = data.split("\"")[1].replaceAll("\\d*$", "");
									int jk = 0;
									for(int jkk = 0;jkk<width.length();jkk++)
									{
										if(jkk == 0 && width.length() == 2){
											jk += (Character.getNumericValue(width.charAt(jkk)) - 9) * 26;
										}else{
											jk += Character.getNumericValue(width.charAt(jkk)) - 9;
										}
									}
									if (!(jk <= lastColumn) && lastColumn != jk-1){
										for(int r = lastColumn;r<jk-1;r++){
											out.write("|");
										}
										
											//out.write("|");
									}
									lastColumn = jk;
									
									if (data.endsWith("/")) {
										out.write("|");
										
									} else if (data.split("\"")[(data.split("\"").length - 1)].equals("s")) {
										isString = 1;
										isData = 1;
									} else {
										if(find(data.split("\"")," s=") == 1 && find(data.split("\"")," s=") != 0 && find(datestyle.split(" "),(data.split("\"")[findvalue(data.split("\"")," s=") + 1])) != 0 && find(datestyle.split(" "),(data.split("\"")[findvalue(data.split("\"")," s=") + 1])) == 1){
											isDate = 1;
										}else{
											isNum = 1;
										}
										isData = 1;
									}
								}
							}
						}
						if (data.startsWith("/row")) {
							inrow = 0;
							out.write("\n");
							needline = 0;
						}
						if (data == "/sheetdata") {
							insheet = 0;
						}
						tag = 0;
						data = null;
					} else {
						if (data == null){
							data = "";
							data += (char)c;
						} else {
							data += (char)c;
						}
					}
				
				}
				System.out.println("Finished : Work Sheet " + (i + 1) );
			}
			
		} finally {
            if (in != null) {
                in.close();
            }
            if (out != null) {
                out.close();
            }
		}
	}

	public static int find(String[] array, String value) {
		for(int i=0; i<array.length; i++) {
			if(array[i].equals(value)){
				return 1;
			}
		}
		return 0;
	}
	public static int findvalue(String[] array, String value) {
		for(int i=0; i<array.length; i++) {
			if(array[i].equals(value)){
				return i;
			}
		}
		return 0;
	}
}