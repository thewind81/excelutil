<%@ page contentType="text/html;charset=UTF-8" language="java" %>
<html>
<head>
	<title>Simple Excel Exporter</title>
</head>
<body>
<h1>Simple Excel Exporter</h1>
<form action="/export">
	<tr>
		<td>Row Count:</td>
		<td><input type="text" name="rowCount" value="1000"></td>
	</tr>
	<tr>
		<td colspan="2"><input type="submit" name="Export Excel"></td>
	</tr>
</form>
</body>
</html>