CREATE OR ALTER PROCEDURE dbo.usp_ScanOrgFiles
AS
BEGIN
    SET NOCOUNT ON;

    DECLARE @Root NVARCHAR(500)  = '\\10.10.1.100\Data\test\';
    DECLARE @cmd  NVARCHAR(1000) = 'dir "' + @Root + '" /A-D /S';

    CREATE TABLE #DirOut (id INT IDENTITY, line NVARCHAR(1000));
    INSERT INTO #DirOut EXEC xp_cmdshell @cmd;

    SELECT
        -- Folder A, B, C etc — first segment after root
        TopFolder  = LEFT(PathPart, CHARINDEX('\', PathPart + '\') - 1),

        -- Everything after the top folder (could be DataType, FileSourceFolder etc)
        SubPath    = CASE
                         WHEN CHARINDEX('\', PathPart) = 0 THEN NULL
                         ELSE SUBSTRING(PathPart, CHARINDEX('\', PathPart) + 1, LEN(PathPart))
                     END,

        FileName   = fn.FileName,
        FilePath   = CurPath + '\' + fn.FileName,
        FileDate   = TRY_CONVERT(DATETIME, LEFT(TRIM(f.line), 10), 101)

    FROM #DirOut f
    CROSS APPLY (
        SELECT TOP 1 CurPath = TRIM(SUBSTRING(d.line, CHARINDEX('Directory of ', d.line) + 13, LEN(d.line)))
        FROM #DirOut d
        WHERE d.id < f.id
          AND d.line LIKE '% Directory of %'
        ORDER BY d.id DESC
    ) dir
    CROSS APPLY (
        SELECT PathPart = SUBSTRING(CurPath, LEN(@Root) + 1, LEN(CurPath))
    ) p
    CROSS APPLY (
        SELECT AfterAMPM = TRIM(SUBSTRING(f.line, PATINDEX('%[AP]M %', f.line) + 2, LEN(f.line)))
    ) t
    CROSS APPLY (
        SELECT FileName = TRIM(SUBSTRING(AfterAMPM, CHARINDEX(' ', AfterAMPM), LEN(AfterAMPM)))
    ) fn
    WHERE TRIM(f.line) LIKE '[0-9][0-9]/[0-9][0-9]/[0-9][0-9][0-9][0-9]%'
      AND PathPart <> ''
    ORDER BY TopFolder, SubPath, FileName;

    DROP TABLE #DirOut;

END
GO

EXEC dbo.usp_ScanOrgFiles;
