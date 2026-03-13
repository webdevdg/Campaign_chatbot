import ReactMarkdown from "react-markdown";
import remarkGfm from "remark-gfm";
import ReportDownload from "./ReportDownload";

function AssistantAvatar() {
  return (
    <div className="w-7 h-7 rounded-full bg-indigo-100 flex items-center justify-center flex-shrink-0 mt-0.5">
      <svg className="w-4 h-4 text-indigo-600" fill="none" viewBox="0 0 24 24" stroke="currentColor" strokeWidth={1.5}>
        <path strokeLinecap="round" strokeLinejoin="round" d="M3 13.125C3 12.504 3.504 12 4.125 12h2.25c.621 0 1.125.504 1.125 1.125v6.75C7.5 20.496 6.996 21 6.375 21h-2.25A1.125 1.125 0 0 1 3 19.875v-6.75ZM9.75 8.625c0-.621.504-1.125 1.125-1.125h2.25c.621 0 1.125.504 1.125 1.125v11.25c0 .621-.504 1.125-1.125 1.125h-2.25a1.125 1.125 0 0 1-1.125-1.125V8.625ZM16.5 4.125c0-.621.504-1.125 1.125-1.125h2.25C20.496 3 21 3.504 21 4.125v15.75c0 .621-.504 1.125-1.125 1.125h-2.25a1.125 1.125 0 0 1-1.125-1.125V4.125Z" />
      </svg>
    </div>
  );
}

export default function MessageBubble({ message }) {
  const isUser = message.role === "user";

  if (isUser) {
    return (
      <div className="flex justify-end mb-6">
        <div className="max-w-[75%] bg-indigo-600 text-white rounded-2xl rounded-br-sm px-4 py-2.5">
          <p className="text-[0.9rem] whitespace-pre-wrap leading-relaxed">{message.content}</p>
        </div>
      </div>
    );
  }

  const displayContent = message.content.replace(/\n*\[Report Config:.*?\]/g, "").trim();

  return (
    <div className="flex gap-3 mb-6">
      <AssistantAvatar />
      <div className="flex-1 min-w-0">
        <div
          className={`assistant-markdown ${
            message.isError ? "text-red-600" : ""
          }`}
        >
          <ReactMarkdown remarkPlugins={[remarkGfm]}>
            {displayContent}
          </ReactMarkdown>
        </div>
        {message.reportUrl && (
          <div className="mt-3">
            <ReportDownload url={message.reportUrl} />
          </div>
        )}
      </div>
    </div>
  );
}
