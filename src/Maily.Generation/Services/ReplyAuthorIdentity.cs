using Maily.Generation.Contracts;

namespace Maily.Generation.Services;

// Incoming-message metadata must never silently become the reply's author.
// Only the compose adapter can mark an explicitly selected outgoing From field.
internal sealed record ReplyAuthorIdentity(string Name, string Address)
{
    public static ReplyAuthorIdentity Resolve(GenerateReplyRequest request)
    {
        var outgoing = request.Compose.FromIsComposeSender &&
            !string.IsNullOrWhiteSpace(request.Compose.FromAddress);
        var name = outgoing ? request.Compose.FromDisplayName : request.Client.MailboxDisplayName;
        var address = outgoing ? request.Compose.FromAddress : request.Client.MailboxAddress;
        return new(string.IsNullOrWhiteSpace(name) ? "Current sender" : name.Trim(), address.Trim());
    }
}
