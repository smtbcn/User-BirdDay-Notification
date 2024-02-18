<%
Dim usersXml, userNode, today, username, userBirthday, remainingDays
Dim approachingBirthdays, birthdaysToday

' Kullanıcıları içeren XML dosyasını yükle
Set usersXml = Server.CreateObject("Microsoft.XMLDOM")
usersXml.Async = False
usersXml.Load(Server.MapPath("assets/birddays/users.xml"))

' Bugünün tarihini al
today = Date()
approachingBirthdays = ""
birthdaysToday = ""

' Her kullanıcı için doğum gününe belirli bir süre kala hatırlatma yap
For Each userNode In usersXml.SelectNodes("/users/user")
    username = userNode.SelectSingleNode("name").Text
    userBirthday = DateSerial(Year(today), Month(userNode.SelectSingleNode("birthday").Text), Day(userNode.SelectSingleNode("birthday").Text))

    remainingDays = DateDiff("d", today, userBirthday)

    ' Doğum gününe 2 gün kala veya bugün doğum günü ise
    If remainingDays >= 0 And remainingDays <= 2 Then
        If remainingDays = 0 Then
            birthdaysToday = birthdaysToday & "<span style='font-weight: bold; color:#000;'>" & username & "</span>" & " Nice Senelere... "
        Else
            approachingBirthdays = approachingBirthdays & "<span style='font-weight: bold; color:#000;'>" & username & "</span>" & " " & remainingDays & " gün kaldı... "
        End If
    End If
Next

%>
<%if birthdaysToday <> "" or approachingBirthdays <> "" then%>
<div class='card bg-info text-white shadow p-3'>
<%
' Bugün doğum günü olanları ve yaklaşan doğum günlerini yazdır
If birthdaysToday <> "" Then
    Response.Write "<span>Bugün doğum günü: " & birthdaysToday & "</span><br>"
End If

If approachingBirthdays <> "" Then
    Response.Write "<span>Yaklaşan doğum günleri: " & approachingBirthdays & "</span>"
End If
%>
</div>
<%else%>

<%End If%>
