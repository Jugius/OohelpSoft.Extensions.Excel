namespace OohelpSoft.Excel;

public readonly struct ProgressState : IEquatable<ProgressState>
{
    public int Total { get; }
    public int Completed { get; }
    public ProgressState(int total, int completed)
    {
        Total = total;
        Completed = completed;
    }

    public override bool Equals(object? obj)
    {
        return obj is ProgressState other && Equals(other);
    }
    public bool Equals(ProgressState other) => this.Total == other.Total && this.Completed == other.Completed;
    public override int GetHashCode()
    {
        return HashCode.Combine(this.Total, this.Completed);
    }

    public static bool operator ==(ProgressState left, ProgressState right)
    {
        return left.Equals(right);
    }

    public static bool operator !=(ProgressState left, ProgressState right)
    {
        return !(left == right);
    }
}
