namespace OohelpSoft.Excel;

public readonly struct ProgressState
{
    public int Total { get; }
    public int Completed { get; }
    public ProgressState(int total, int completed)
    {
        Total = total;
        Completed = completed;
    }
}
