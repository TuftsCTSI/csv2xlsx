test: csv2xlsx.jar
	echo "a,b\n1,2" | java -jar csv2xlsx.jar

csv2xlsx.jar: csv2xlsx.java
	mvn package

clean:
	rm -rf *.xlsx target

gitadd:
	git add src/main/java/org/tuftsctsi/CSV2XLSX.java
