0|1|2|3|4|5|6|7
</br>// Gán giá trị vào biến</br>
O: id = parts[0].trim();</br>
1: if (parts.length > 0) email = parts[1].trim();</br>
2: if (parts.length > 1) password = parts[2].trim();</br>
3: if (parts.length > 2) expiryDate = Integer.valueOf(parts[3].trim());</br>
4: if (parts.length > 3) extra = parts[4].trim();</br>
5: if (parts.length > 4) receiveEmail = parts[5].trim();</br>
6: if (parts.length > 5) sendLogFlag = parts[6].trim().equals("1");</br>
7: if (parts.length > 6) hasUpdate = parts[7].trim().equals("1");</br>
