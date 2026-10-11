namespace SharpCell.Tests.Model;

/// <summary>A sheet can be protected, with or without a password; the password is kept only as Excel's hash.</summary>
public class ProtectionTests
{
    private static Worksheet NewSheet() => new Workbook().AddSheet("S");

    [Fact]
    public void A_sheet_is_unprotected_until_protected()
    {
        var sheet = NewSheet();

        Assert.Null(sheet.Protection);
        Assert.Null(sheet.ProtectionPassword);
    }

    [Fact]
    public void Protecting_with_a_password_keeps_its_hash_and_allows_the_defaults()
    {
        var sheet = NewSheet();

        sheet.Protect("secret");

        Assert.Equal(new SheetProtection(), sheet.Protection);
        var password = sheet.ProtectionPassword!;
        Assert.True(password.Verify("secret"));
        Assert.False(password.Verify("Secret"));
        Assert.False(password.Verify(""));
    }

    [Fact]
    public void Protecting_without_a_password_keeps_none()
    {
        var sheet = NewSheet();

        sheet.Protect();
        Assert.Equal(new SheetProtection(), sheet.Protection);
        Assert.Null(sheet.ProtectionPassword);

        sheet.Protect("");
        Assert.Null(sheet.ProtectionPassword);
    }

    [Fact]
    public void Protecting_with_allowed_actions_keeps_them()
    {
        var sheet = NewSheet();
        var allow = new SheetProtection { FormatCells = true, Sort = true, SelectLockedCells = false };

        sheet.Protect("secret", allow);

        Assert.Same(allow, sheet.Protection);
    }

    [Fact]
    public void Unprotecting_removes_the_protection_and_the_password()
    {
        var sheet = NewSheet();
        sheet.Protect("secret");

        sheet.Unprotect();

        Assert.Null(sheet.Protection);
        Assert.Null(sheet.ProtectionPassword);
    }

    [Fact]
    public void Protection_does_not_stop_code_from_changing_the_sheet()
    {
        var sheet = NewSheet();
        sheet.Protect("secret");

        sheet["A1"].Value = 5;
        sheet.InsertRows(1);

        Assert.Equal(CellValue.Number(5), sheet["A2"].Value);
    }

    [Fact]
    public void The_defaults_allow_selecting_cells_and_nothing_else()
    {
        var allow = new SheetProtection();

        Assert.True(allow.SelectLockedCells);
        Assert.True(allow.SelectUnlockedCells);
        Assert.False(allow.FormatCells || allow.FormatColumns || allow.FormatRows || allow.InsertColumns || allow.InsertRows
            || allow.InsertHyperlinks || allow.DeleteColumns || allow.DeleteRows || allow.Sort || allow.AutoFilter
            || allow.PivotTables || allow.EditObjects || allow.EditScenarios);
        Assert.Equal("SelectLockedCells, SelectUnlockedCells", allow.ToString());
        Assert.Equal("None", new SheetProtection { SelectLockedCells = false, SelectUnlockedCells = false }.ToString());
    }

    [Fact]
    public void Protections_allowing_the_same_actions_are_equal()
    {
        var all = new SheetProtection
        {
            FormatCells = true, FormatColumns = true, FormatRows = true, InsertColumns = true, InsertRows = true,
            InsertHyperlinks = true, DeleteColumns = true, DeleteRows = true, Sort = true, AutoFilter = true,
            PivotTables = true, EditObjects = true, EditScenarios = true,
        };

        Assert.Equal(new SheetProtection { Sort = true }, new SheetProtection { Sort = true });
        Assert.Equal(new SheetProtection { Sort = true }.GetHashCode(), new SheetProtection { Sort = true }.GetHashCode());
        Assert.NotEqual(new SheetProtection(), new SheetProtection { Sort = true });
        Assert.NotEqual(new SheetProtection(), new SheetProtection { SelectUnlockedCells = false });
        Assert.Equal(all, new SheetProtection(all));
        Assert.NotEqual(all, new SheetProtection(all) { EditScenarios = false });
        Assert.False(new SheetProtection().Equals(null));
        Assert.Throws<ArgumentNullException>(() => new SheetProtection(null!));
    }

    [Fact]
    public void A_new_password_is_hashed_as_Excel_hashes_it()
    {
        var password = ProtectionPassword.Create("secret");
        var other = ProtectionPassword.Create("secret");

        Assert.Equal("SHA-512", password.AlgorithmName);
        Assert.Equal(100_000, password.SpinCount);
        Assert.Equal(16, Convert.FromBase64String(password.SaltValue!).Length);
        Assert.Equal(64, Convert.FromBase64String(password.HashValue!).Length);
        Assert.Null(password.LegacyHash);
        Assert.NotEqual(password.SaltValue, other.SaltValue);   // a random salt each time
        Assert.True(other.Verify("secret"));
    }

    [Fact]
    public void A_hash_is_checked_with_the_salt_and_spin_count_it_was_made_with()
    {
        var password = ProtectionPassword.Create("secret");

        Assert.False((password with { SpinCount = 99_999 }).Verify("secret"));
        Assert.False((password with { SaltValue = ProtectionPassword.Create("secret").SaltValue }).Verify("secret"));
        Assert.False((password with { AlgorithmName = "SHA-1" }).Verify("secret"));
    }

    [Fact]
    public void A_legacy_password_is_kept_but_not_verified()
    {
        var password = new ProtectionPassword { LegacyHash = "CC3D" };

        Assert.False(password.Verify("secret"));
    }

    [Fact]
    public void A_clone_keeps_the_protection_and_its_password()
    {
        var workbook = new Workbook();
        var sheet = workbook.AddSheet("S");
        sheet.Protect("secret", new SheetProtection { Sort = true });

        var clone = workbook.Clone();

        Assert.Equal(new SheetProtection { Sort = true }, clone["S"].Protection);
        Assert.True(clone["S"].ProtectionPassword!.Verify("secret"));
        clone["S"].Unprotect();
        Assert.NotNull(sheet.Protection);
    }

    [Fact]
    public void Protecting_from_a_custom_function_is_refused()
    {
        var workbook = new Workbook();
        var sheet = workbook.AddSheet("S");
        var caught = new List<Exception>();
        workbook.Functions.Add("HOOK", _ =>
        {
            try { sheet.Protect("secret"); } catch (Exception ex) { caught.Add(ex); }
            try { sheet.Unprotect(); } catch (Exception ex) { caught.Add(ex); }
            return 5;
        });
        sheet.Protect();
        sheet["A1"].Formula = "=HOOK()";
        workbook.Recalculate();

        Assert.Equal(2, caught.Count);
        Assert.All(caught, ex => Assert.IsType<InvalidOperationException>(ex));
        Assert.NotNull(sheet.Protection);
        Assert.Null(sheet.ProtectionPassword);
    }

    [Fact]
    public void A_workbook_is_unprotected_until_protected()
    {
        var workbook = new Workbook();

        Assert.False(workbook.IsStructureProtected);
        Assert.False(workbook.IsWindowsProtected);
        Assert.Null(workbook.ProtectionPassword);
    }

    [Fact]
    public void Protecting_a_workbook_locks_its_structure_and_hashes_the_password()
    {
        var workbook = new Workbook();

        workbook.Protect("secret");

        Assert.True(workbook.IsStructureProtected);
        Assert.False(workbook.IsWindowsProtected);
        Assert.True(workbook.ProtectionPassword!.Verify("secret"));
        Assert.False(workbook.ProtectionPassword.Verify("Secret"));
    }

    [Fact]
    public void A_workbook_can_lock_its_windows_instead_of_its_structure()
    {
        var workbook = new Workbook();

        workbook.Protect(structure: false, windows: true);

        Assert.False(workbook.IsStructureProtected);
        Assert.True(workbook.IsWindowsProtected);
        Assert.Null(workbook.ProtectionPassword);
        Assert.Throws<ArgumentException>(() => workbook.Protect(structure: false, windows: false));
    }

    [Fact]
    public void Unprotecting_a_workbook_removes_the_locks_and_the_password()
    {
        var workbook = new Workbook();
        workbook.Protect("secret", windows: true);

        workbook.Unprotect();

        Assert.False(workbook.IsStructureProtected);
        Assert.False(workbook.IsWindowsProtected);
        Assert.Null(workbook.ProtectionPassword);
    }

    [Fact]
    public void A_clone_keeps_the_workbook_protection()
    {
        var workbook = new Workbook();
        workbook.Protect("secret", windows: true);

        var clone = workbook.Clone();

        Assert.True(clone.IsStructureProtected);
        Assert.True(clone.IsWindowsProtected);
        Assert.True(clone.ProtectionPassword!.Verify("secret"));
    }
}
