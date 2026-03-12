import useChat from "./hooks/useChat";
import ChatWindow from "./components/ChatWindow";

export default function App() {
  const { messages, isLoading, sendMessage } = useChat();

  return (
    <div className="h-full bg-white flex flex-col">
      <header className="border-b border-gray-200 px-6 py-4 flex-shrink-0">
        <div className="max-w-3xl mx-auto">
          <h1 className="text-lg font-semibold text-gray-900">Campaign Analyst</h1>
          <p className="text-sm text-gray-400 mt-0.5">
            Ask anything about your Q1 2026 media campaign performance
          </p>
        </div>
      </header>

      <main className="flex-1 overflow-hidden max-w-3xl w-full mx-auto">
        <ChatWindow
          messages={messages}
          isLoading={isLoading}
          onSend={sendMessage}
        />
      </main>
    </div>
  );
}
