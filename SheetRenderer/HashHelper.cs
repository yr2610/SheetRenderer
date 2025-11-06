using System;
using System.IO;
using System.Security.Cryptography;
using System.Text;

internal static class HashHelper
{
    // 共通: ハッシュ計算
    private static string Compute(HashAlgorithm algo, string text)
    {
        if (text == null) text = string.Empty;
        var bytes = Encoding.UTF8.GetBytes(text);
        var hash = algo.ComputeHash(bytes);
        return ToHex(hash);
    }

    // 16進文字列に変換（小文字）
    public static string ToHex(byte[] bytes)
    {
        var sb = new StringBuilder(bytes.Length * 2);
        for (int i = 0; i < bytes.Length; i++)
        {
            sb.AppendFormat("{0:x2}", bytes[i]);
        }
        return sb.ToString();
    }

    // ---- MD5 ----
    public static string Md5(string text)
    {
        using (var md5 = MD5.Create())
        {
            return Compute(md5, text);
        }
    }

    // ---- SHA1 ----
    public static string Sha1(string text)
    {
        using (var sha1 = SHA1.Create())
        {
            return Compute(sha1, text);
        }
    }

    // ---- SHA256 ----
    public static string Sha256(string text)
    {
        using (var sha256 = SHA256.Create())
        {
            return Compute(sha256, text);
        }
    }

    // ---- ファイル版（SHA256例）----
    public static string Sha256File(string path)
    {
        using (var sha256 = SHA256.Create())
        using (var fs = File.OpenRead(path))
        {
            var hash = sha256.ComputeHash(fs);
            return ToHex(hash);
        }
    }
}

// JS に公開する橋渡し（インスタンス）
public sealed class HashBridge
{
    public string md5(string text) { return HashHelper.Md5(text); }
    public string sha1(string text) { return HashHelper.Sha1(text); }
    public string sha256(string text) { return HashHelper.Sha256(text); }
}
