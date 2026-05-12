namespace Converter;

public interface IFileToMarkdownConverter
{
    // sourcePath: the file actually being read (may be a temp file produced by an upgrader).
    // displaySourcePath: the path the user sees, used as the prefix for the .ex error file.
    //                    For native formats this equals sourcePath.
    // outputPath: where the resulting <name>.<ext>.md should land.
    void Convert(string sourcePath, string displaySourcePath, string outputPath);
}
