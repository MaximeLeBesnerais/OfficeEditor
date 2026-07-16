interface StatusMessageProps {
  type: 'error' | 'warning' | 'success' | 'info';
  title?: string;
  message: string;
  messages?: string[];
}

const styles: Record<StatusMessageProps['type'], string> = {
  error: 'bg-red-50 text-red-900 border-red-200',
  warning: 'bg-amber-50 text-amber-900 border-amber-200',
  success: 'bg-emerald-50 text-emerald-900 border-emerald-200',
  info: 'bg-blue-50 text-blue-900 border-blue-200',
};

const iconBg: Record<StatusMessageProps['type'], string> = {
  error: 'bg-red-100 text-red-600',
  warning: 'bg-amber-100 text-amber-600',
  success: 'bg-emerald-100 text-emerald-600',
  info: 'bg-blue-100 text-blue-600',
};

const paths: Record<StatusMessageProps['type'], string> = {
  error: 'M12 9v3.75m-9.303 3.376c-.866 1.5.217 3.374 1.948 3.374h14.71c1.73 0 2.813-1.874 1.948-3.374L13.949 3.378c-.866-1.5-3.032-1.5-3.898 0L2.697 16.126ZM12 15.75h.007v.008H12v-.008Z',
  warning:
    'M12 9v3.75m-9.303 3.376c-.866 1.5.217 3.374 1.948 3.374h14.71c1.73 0 2.813-1.874 1.948-3.374L13.949 3.378c-.866-1.5-3.032-1.5-3.898 0L2.697 16.126ZM12 15.75h.007v.008H12v-.008Z',
  success:
    'M9 12.75 11.25 15 15 9.75M21 12a9 9 0 1 1-18 0 9 9 0 0 1 18 0Z',
  info: 'M11.25 11.25l.041-.02a.75.75 0 0 1 1.063.852l-.708 2.836a.75.75 0 0 0 1.063.853l.041-.021M21 12a9 9 0 1 1-18 0 9 9 0 0 1 18 0ZM8.25 9.75h.008v.008H8.25V9.75Zm6 0h.008v.008H14.25V9.75Z',
};

export function StatusMessage({ type, title, message, messages }: StatusMessageProps) {
  return (
    <div
      className={[
        'rounded-2xl border p-4 shadow-sm',
        'flex items-start gap-3',
        styles[type],
      ].join(' ')}
      role={type === 'error' ? 'alert' : 'status'}
    >
      <div
        className={[
          'w-9 h-9 rounded-lg flex items-center justify-center shrink-0',
          iconBg[type],
        ].join(' ')}
      >
        <svg
          xmlns="http://www.w3.org/2000/svg"
          fill="none"
          viewBox="0 0 24 24"
          strokeWidth={1.5}
          stroke="currentColor"
          className="w-5 h-5"
          aria-hidden="true"
        >
          <path strokeLinecap="round" strokeLinejoin="round" d={paths[type]} />
        </svg>
      </div>
      <div className="min-w-0">
        {title && <p className="font-semibold">{title}</p>}
        <p className="text-sm leading-relaxed">{message}</p>
        {messages && messages.length > 0 && (
          <ul className="mt-2 text-sm list-disc list-inside opacity-90">
            {messages.map((item, index) => (
              <li key={index}>{item}</li>
            ))}
          </ul>
        )}
      </div>
    </div>
  );
}
