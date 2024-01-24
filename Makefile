csv2xlsx: csv2xlsx.jar
	java -jar csv2xlsx.jar

csv2xlsx.jar: csv2xlsx.java
	mvn package

gitadd:
	git add src/main/java/org/tuftsctsi/CSV2XLSX.java
