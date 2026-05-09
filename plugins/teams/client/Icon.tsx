import * as React from "react";

type Props = {
  /** The size of the icon, 24px is default to match standard icons */
  size?: number;
  /** The color of the icon, defaults to the current text color */
  fill?: string;
  /** Whether to render the monochrome version, defaults to true */
  monochrome?: boolean;
};

export default function Icon({
  size = 24,
  fill = "currentColor",
  monochrome = true,
}: Props) {
  if (monochrome) {
    return (
      <svg
        fill={fill}
        width={size}
        height={size}
        viewBox="0 0 24 24"
        version="1.1"
      >
        <path d="M20.625 5.8125H13.125C12.7109 5.8125 12.375 6.14844 12.375 6.5625V8.4375H9.375C8.96094 8.4375 8.625 8.77344 8.625 9.1875V18.1875C8.625 18.6016 8.96094 18.9375 9.375 18.9375H16.875C17.2891 18.9375 17.625 18.6016 17.625 18.1875V16.3125H20.625C21.0391 16.3125 21.375 15.9766 21.375 15.5625V6.5625C21.375 6.14844 21.0391 5.8125 20.625 5.8125ZM16.125 17.4375H10.125V9.9375H16.125V17.4375ZM19.875 14.8125H17.625V9.1875C17.625 8.77344 17.2891 8.4375 16.875 8.4375H13.875V7.3125H19.875V14.8125ZM10.875 12.1875H15.375V13.125H10.875V12.1875ZM10.875 14.0625H15.375V15H10.875V14.0625ZM10.875 10.3125H15.375V11.25H10.875V10.3125ZM5.25 4.875C4.42969 4.875 3.75 5.55469 3.75 6.375V13.125C3.75 13.9453 4.42969 14.625 5.25 14.625C6.07031 14.625 6.75 13.9453 6.75 13.125V6.375C6.75 5.55469 6.07031 4.875 5.25 4.875ZM5.25 13.5C5.04688 13.5 4.875 13.3281 4.875 13.125V11.8125H5.625V13.125C5.625 13.3281 5.45313 13.5 5.25 13.5ZM5.625 10.6875H4.875V9H5.625V10.6875ZM5.625 7.875H4.875V6.375C4.875 6.17188 5.04688 6 5.25 6C5.45313 6 5.625 6.17188 5.625 6.375V7.875Z" />
      </svg>
    );
  }

  return (
    <svg
      width={size}
      height={size}
      viewBox="0 0 24 24"
      version="1.1"
      fill="none"
    >
      <path
        d="M14.25 7.5H21.75V15H17.25V18.75H9.75V11.25H14.25V7.5Z"
        fill="#5059C9"
      />
      <path
        d="M9.75 11.25H17.25V18.75H9.75V11.25Z"
        fill="#7B83EB"
      />
      <path
        d="M9.75 11.25H17.25V13.5H9.75V11.25Z"
        fill="#000000"
        fillOpacity="0.1"
      />
      <circle cx="5.625" cy="7.875" r="2.625" fill="#5059C9" />
      <path
        d="M3 11.25H8.25V16.5H5.625C4.17525 16.5 3 15.3248 3 13.875V11.25Z"
        fill="#7B83EB"
      />
    </svg>
  );
}
