// Copyright 2026 OfficeCLI (https://OfficeCLI.AI)
// SPDX-License-Identifier: Apache-2.0

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using W15 = DocumentFormat.OpenXml.Office2013.Word;

namespace OfficeCli.Handlers;

public partial class WordHandler
{
    // Modern comment metadata lives in word/commentsExtended.xml (w15) — reply
    // threading (w15:paraIdParent) and resolved-state (w15:done), keyed by each
    // comment's first-paragraph w14:paraId (NOT by w:id). The user-facing CLI
    // speaks comment ids (the /comments/comment[@commentId=N] the user sees);
    // these helpers translate id <-> paraId and own the w15 part so AddComment /
    // SetElementComment / readback never duplicate the part plumbing.

    /// <summary>Get-or-create the CommentsEx root in word/commentsExtended.xml.</summary>
    private W15.CommentsEx EnsureCommentsExRoot()
    {
        var main = _doc.MainDocumentPart
            ?? throw new InvalidOperationException("Document main part not found");
        var part = main.WordprocessingCommentsExPart
            ?? main.AddNewPart<WordprocessingCommentsExPart>();
        part.CommentsEx ??= new W15.CommentsEx();
        return part.CommentsEx;
    }

    /// <summary>
    /// First-paragraph w14:paraId of the comment with the given w:id, assigning a
    /// fresh paraId if the paragraph lacks one. Returns null when no such comment.
    /// </summary>
    private string? GetCommentFirstParaId(string commentId)
    {
        var comment = _doc.MainDocumentPart?.WordprocessingCommentsPart?.Comments?
            .Elements<Comment>().FirstOrDefault(c => c.Id?.Value == commentId);
        var firstPara = comment?.Descendants<Paragraph>().FirstOrDefault();
        if (firstPara == null) return null;
        if (string.IsNullOrEmpty(firstPara.ParagraphId?.Value)) AssignParaId(firstPara);
        return firstPara.ParagraphId?.Value;
    }

    /// <summary>
    /// Find-or-create the w15:commentEx for <paramref name="paraId"/> and apply the
    /// supplied parent/done (null = leave unchanged). New entries default Done=false
    /// to mirror Word's <c>w15:done="0"</c> output. Saves the part.
    /// </summary>
    private void UpsertCommentEx(string paraId, string? parentParaId, bool? done)
    {
        var root = EnsureCommentsExRoot();
        var ex = root.Elements<W15.CommentEx>().FirstOrDefault(e => e.ParaId?.Value == paraId);
        if (ex == null)
        {
            ex = new W15.CommentEx { ParaId = paraId, Done = OnOffValue.FromBoolean(false) };
            root.AppendChild(ex);
        }
        if (parentParaId != null) ex.ParaIdParent = parentParaId;
        if (done.HasValue) ex.Done = OnOffValue.FromBoolean(done.Value);
        _doc.MainDocumentPart!.WordprocessingCommentsExPart!.CommentsEx!.Save();
    }

    /// <summary>
    /// Read the resolved-state and reply-parent of a comment for Get/Query.
    /// Returns the PARENT COMMENT's w:id (translated back from w15:paraIdParent)
    /// so the readback matches the id the caller passes to `--prop parentId=`,
    /// and the done flag (false when the comment has no commentEx entry).
    /// </summary>
    private (string? parentId, bool done) ReadCommentExInfo(Comment comment)
    {
        var paraId = comment.Descendants<Paragraph>().FirstOrDefault()?.ParagraphId?.Value;
        var part = _doc.MainDocumentPart?.WordprocessingCommentsExPart;
        if (string.IsNullOrEmpty(paraId) || part?.CommentsEx == null) return (null, false);
        var ex = part.CommentsEx.Elements<W15.CommentEx>()
            .FirstOrDefault(e => e.ParaId?.Value == paraId);
        if (ex == null) return (null, false);
        bool done = TryReadOnOff(ex.Done) == true;
        string? parentId = null;
        var parentParaId = ex.ParaIdParent?.Value;
        if (!string.IsNullOrEmpty(parentParaId))
        {
            var parent = _doc.MainDocumentPart?.WordprocessingCommentsPart?.Comments?
                .Elements<Comment>().FirstOrDefault(c =>
                    c.Descendants<Paragraph>().FirstOrDefault()?.ParagraphId?.Value == parentParaId);
            parentId = parent?.Id?.Value;
        }
        return (parentId, done);
    }
}
