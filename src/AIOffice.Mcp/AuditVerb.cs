using AIOffice.Core;

namespace AIOffice.Mcp;

/// <summary>
/// The shared audit pipeline behind both <c>aioffice audit</c> and the
/// <c>office_audit</c> MCP tool, so the two surfaces return a byte-identical
/// data payload. Auditing is read-only by default; with <c>--fix</c> / <c>fix:true</c>
/// it applies the safe autofix subset, then re-audits so the caller sees what
/// remains. Findings are data, not errors — a successful audit always reports
/// ok:true (exit 0), even when it surfaces error-severity findings.
/// </summary>
public static class AuditVerb
{
    /// <summary>
    /// Runs the audit (and, when <see cref="AuditOptions.Fix"/>, the autofix pass)
    /// and shapes the result envelope:
    /// <c>{findings:[…], summary:{errors,warnings,infos}, fixed?:N, remaining?:[…]}</c>.
    /// </summary>
    public static Envelope Run(IAuditor auditor, CommandContext ctx, AuditOptions opts)
    {
        if (!opts.Fix)
        {
            var result = auditor.Audit(ctx, opts);
            return Envelope.Ok(new
            {
                findings = result.Findings,
                summary = result.Summary,
            });
        }

        // --fix: audit first so we know which findings are autofixable, then
        // apply the safe subset (empty id list = every autofixable finding),
        // then re-audit so the caller sees exactly what is left.
        var before = auditor.Audit(ctx, opts);
        var fixedCount = auditor.Fix(ctx, []);
        var after = auditor.Audit(ctx, opts);

        return Envelope.Ok(new
        {
            findings = after.Findings,
            summary = after.Summary,
            @fixed = fixedCount,
            remaining = after.Findings.Select(f => f.Id).ToList(),
        });
    }
}
