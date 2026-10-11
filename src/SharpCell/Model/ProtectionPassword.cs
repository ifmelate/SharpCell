using System;
using System.Buffers.Binary;
using System.Security.Cryptography;
using System.Text;

namespace SharpCell;

/// <summary>
/// The password of a protected sheet (or workbook) as files keep it: a salted, iterated hash, or a
/// legacy 16-bit one. The plain password is never kept. A password read from a file is held as the
/// file has it, so it is written back unchanged; which attributes hold it is the file format's business.
/// </summary>
internal sealed record ProtectionPassword
{
    private const string Sha512 = "SHA-512";
    private const int SaltLength = 16;
    private const int DefaultSpinCount = 100_000;

    /// <summary>The hash algorithm's name as files write it, such as <c>SHA-512</c>; null for a legacy password.</summary>
    public string? AlgorithmName { get; init; }

    /// <summary>The hash, in base64.</summary>
    public string? HashValue { get; init; }

    /// <summary>The salt, in base64.</summary>
    public string? SaltValue { get; init; }

    /// <summary>How many times the hash was hashed again.</summary>
    public int? SpinCount { get; init; }

    /// <summary>The legacy 16-bit hash in hexadecimal, as older Excel versions wrote it; null for none.</summary>
    public string? LegacyHash { get; init; }

    /// <summary>The password hashed as Excel hashes it: SHA-512, a random 16-byte salt, 100,000 spins.</summary>
    public static ProtectionPassword Create(string password)
    {
        var salt = RandomNumberGenerator.GetBytes(SaltLength);
        return new ProtectionPassword
        {
            AlgorithmName = Sha512,
            HashValue = Convert.ToBase64String(Hash(password, salt, DefaultSpinCount)),
            SaltValue = Convert.ToBase64String(salt),
            SpinCount = DefaultSpinCount,
        };
    }

    /// <summary>
    /// Whether the password hashes to this one. Only SHA-512 hashes are checked: a legacy 16-bit
    /// hash, another algorithm or a damaged hash gives false.
    /// </summary>
    public bool Verify(string password)
    {
        if (!string.Equals(AlgorithmName, Sha512, StringComparison.OrdinalIgnoreCase)
            || HashValue is null || SaltValue is null || SpinCount is not >= 0)
            return false;
        try
        {
            return CryptographicOperations.FixedTimeEquals(Hash(password, Convert.FromBase64String(SaltValue), SpinCount.Value),
                Convert.FromBase64String(HashValue));
        }
        catch (FormatException)
        {
            return false;
        }
    }

    // ECMA-376 Part 4, 2.3.7.4 (as used for sheetProtection): H0 = H(salt + UTF-16LE password),
    // then Hn = H(Hn-1 + n-1 as a little-endian 32-bit integer) spinCount times.
    private static byte[] Hash(string password, byte[] salt, int spinCount)
    {
        var hash = SHA512.HashData([.. salt, .. Encoding.Unicode.GetBytes(password)]);
        var buffer = new byte[SHA512.HashSizeInBytes + sizeof(int)];
        for (var i = 0; i < spinCount; i++)
        {
            hash.CopyTo(buffer, 0);
            BinaryPrimitives.WriteInt32LittleEndian(buffer.AsSpan(SHA512.HashSizeInBytes), i);
            SHA512.HashData(buffer, hash);
        }

        return hash;
    }
}
