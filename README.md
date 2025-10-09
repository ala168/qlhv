0|1|2|3|4|5|6|7|8|9|10
</br>
// Gán giá trị vào biến
id = parts[0].trim();
if (parts.length > 0) email = parts[1].trim();
if (parts.length > 1) password = parts[2].trim();
if (parts.length > 2) expiryDate = Integer.valueOf(parts[3].trim());
if (parts.length > 3) extra = parts[4].trim();
if (parts.length > 4) receiveEmail = parts[5].trim();
if (parts.length > 5) sendLogFlag = parts[6].trim().equals("1");
if (parts.length > 6) hasUpdate = parts[7].trim().equals("1");
