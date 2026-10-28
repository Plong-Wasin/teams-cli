# Teams CLI

A command-line client for Microsoft Teams channels. It talks to the Microsoft Graph API so people and agents can read and post Channel Messages from a shell, replacing the Teams MCP server.

## Language

**Team**:
A Microsoft Teams team the signed-in user has joined, backed by a Microsoft 365 group. A Team contains Channels.
_Avoid_: group, workspace

**Channel**:
A named conversation space within a Team where Channel Messages are posted and read.
_Avoid_: room

**Channel Message**:
A message posted to a Channel, visible to every Team member with access to it. Every Channel Message is either a Root Message or a Reply.
_Avoid_: post, chat message

**Root Message**:
A Channel Message with no parent; it opens a Thread. Only a Root Message can carry a subject.
_Avoid_: top-level message, parent message

**Reply**:
A Channel Message posted under a Root Message. A Reply cannot itself be replied to; a Thread is one level deep.
_Avoid_: comment, response

**Thread**:
A Root Message together with all of its Replies.
_Avoid_: conversation, discussion

**Chat** _(out of scope)_:
A 1:1 or group conversation outside any Team. No command targets a Chat. The exception is `search`, which queries Microsoft Search across Channel Messages and Chat messages because the API does not separate them: reading a Chat message that way is tolerated, writing to a Chat is not.

**Profile**:
A named set of sign-in and defaults selected with `-p`.
_Avoid_: account, environment, context
