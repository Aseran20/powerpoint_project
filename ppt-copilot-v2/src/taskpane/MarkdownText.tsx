import * as React from "react";
import { parseMarkdown, type FormattedParagraph } from "./markdownParser";

interface MarkdownTextProps {
    text: string;
}

/**
 * Component that renders Markdown text with proper HTML formatting
 */
export const MarkdownText: React.FC<MarkdownTextProps> = ({ text }) => {
    const paragraphs = parseMarkdown(text);

    return (
        <div style={{ whiteSpace: "pre-wrap" }}>
            {paragraphs.map((para, idx) => {
                if (para.isBullet) {
                    // Render as list item
                    return (
                        <div key={idx} style={{ display: "flex", marginBottom: "4px" }}>
                            <span style={{ marginRight: "8px" }}>•</span>
                            <span>{renderSegments(para)}</span>
                        </div>
                    );
                } else {
                    // Render as regular paragraph
                    return (
                        <p key={idx} style={{ margin: "0 0 8px 0" }}>
                            {renderSegments(para)}
                        </p>
                    );
                }
            })}
        </div>
    );
};

/**
 * Render text segments with inline formatting
 */
function renderSegments(para: FormattedParagraph) {
    return para.segments.map((segment, idx) => {
        let content: React.ReactNode = segment.text;

        if (segment.bold && segment.italic) {
            return <strong key={idx}><em>{content}</em></strong>;
        } else if (segment.bold) {
            return <strong key={idx}>{content}</strong>;
        } else if (segment.italic) {
            return <em key={idx}>{content}</em>;
        } else {
            return <span key={idx}>{content}</span>;
        }
    });
}
