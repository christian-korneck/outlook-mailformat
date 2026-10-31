// Run via build.cmd. Exit code 1 on failure.
using System;
using MailformatAddin;

static class Tests
{
    static int failed;

    static void Eq(string name, string actual, string expected)
    {
        if (actual == expected) { Console.WriteLine("ok   " + name); return; }
        failed++;
        Console.WriteLine("FAIL " + name + "\n  expected: " + Show(expected) + "\n  actual:   " + Show(actual));
    }

    static string Show(string s) { return "\"" + s.Replace("\n", "\\n") + "\""; }

    static int Main()
    {
        Eq("wrap short line unchanged", Mailformat.Wrap("hello world", 72), "hello world");
        Eq("wrap greedy", Mailformat.Wrap("aaa bbb ccc ddd", 7), "aaa bbb\nccc ddd");
        Eq("wrap collapses whitespace", Mailformat.Wrap("a   b\tc", 72), "a b c");
        Eq("wrap keeps line breaks", Mailformat.Wrap("one two\nthree", 3), "one\ntwo\nthree");
        Eq("wrap keeps empty lines", Mailformat.Wrap("a\n\nb", 72), "a\n\nb");
        Eq("wrap long word", Mailformat.Wrap("x abcdefghij y", 5), "x\nabcdefghij\ny");
        Eq("wrap keeps quote prefix", Mailformat.Wrap("> aaa bbb ccc", 9), "> aaa bbb\n> ccc");
        Eq("wrap nested prefix", Mailformat.Wrap(">> > aa bb cc", 10), ">> > aa bb\n>> > cc");
        Eq("wrap prefix-only line", Mailformat.Wrap("> ", 72), "> ");
        Eq("wrap width below prefix", Mailformat.Wrap(">>>> a b", 2), ">>>> a\n>>>> b");

        Eq("quote", Mailformat.Quote("a\n\n> b\nc"), "> a\n>\n>> b\n> c");
        Eq("unquote", Mailformat.Unquote("> a\n>\n>> b\nc\n>x"), "a\n\n> b\nc\nx");
        Eq("quote/unquote roundtrip", Mailformat.Unquote(Mailformat.Quote("a\n> b")), "a\n> b");

        Eq("both stays within width", Mailformat.Apply("both", "aaa bbb ccc", 9), "> aaa bbb\n> ccc");
        Eq("both on quoted text", Mailformat.Apply("both", "> aa bb cc", 9), ">> aa bb\n>> cc");

        Console.WriteLine(failed == 0 ? "All tests passed." : failed + " test(s) failed.");
        return failed == 0 ? 0 : 1;
    }
}
