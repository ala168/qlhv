# Đầu vào: JAR gốc
-injars E:\workspace\profile_workspace\QLHV\Run\QLHV.jar

# Đầu ra: JAR sau khi obfuscate
-outjars myapp-obfuscated.jar

# Thư viện Java chuẩn (cần giữ nguyên)
-libraryjars <java.home>/lib/rt.jar

# Giữ lại class main (không obfuscate tên này)
-keep public class HVMain.Main {
    public static void main(java.lang.String[]);
}

# Giữ lại các class dùng reflection (nếu có)
# -keep class com.yourpackage.** { *; }

# Tối ưu hóa và loại bỏ code thừa
-dontoptimize
-dontshrink
-overloadaggressively
