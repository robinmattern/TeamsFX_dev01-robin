import { TeamsFx } from "@microsoft/teamsfx";

export class MyApp extends TeamsFx {

  async getMessages(conversation1Id, conversation2Id) {
    const conversation1 = await this.conversations.get(conversation1Id);
    const conversation2 = await this.conversations.get(conversation2Id);

    const messages = [];
    messages.push(...conversation1.messages);
    messages.push(...conversation2.messages);

    return messages;
  }

  async sendMergedConversation(messages) {
    const conversation = await this.conversations.create();
    conversation.messages = messages;

    await conversation.send();
  }

}