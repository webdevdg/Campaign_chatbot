import useChat from "./hooks/useChat";
import ChatWindow from "./components/ChatWindow";

export default function App() {
  const { messages, isLoading, sendMessage, resetChat } = useChat();

  return (
    <div className="h-full bg-white flex flex-col">
      <header className="border-b border-gray-200 px-6 py-3 flex-shrink-0">
        <div className="max-w-3xl mx-auto flex items-center justify-between">
          <div>
            <h1 className="text-lg font-semibold text-gray-900">Campaign Analyst</h1>
            <p className="text-sm text-gray-400 mt-0.5">
              Ask anything about your Q1 2026 media campaign performance
            </p>
          </div>
          <button
            onClick={resetChat}
            disabled={isLoading || messages.length === 0}
            className="flex items-center gap-1.5 px-3 py-2 text-sm font-medium text-indigo-700 bg-indigo-50 border border-indigo-200 rounded-lg hover:bg-indigo-100 hover:border-indigo-300 disabled:bg-gray-100 disabled:border-gray-200 disabled:text-gray-400 disabled:cursor-not-allowed transition-colors"
            title="Start a new conversation"
          >
            <svg className="w-4 h-4" fill="none" viewBox="0 0 24 24" stroke="currentColor" strokeWidth={1.5}>
              <path strokeLinecap="round" strokeLinejoin="round" d="M12 4.5v15m7.5-7.5h-15" />
            </svg>
            <span className="hidden sm:inline">New Chat</span>
          </button>
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
