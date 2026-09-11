"""Copy the generated plain text on the user's click, without a download."""
from streamlit.components.v2 import component


_copy = component(
    "copy_summary",
    html='''<button type="button"></button><div role="status" aria-live="polite"></div>
<textarea hidden readonly aria-label="직접 복사할 개요" rows="5"></textarea>''',
    css='''
button {width:100%; padding:.65rem 1rem; border:1px solid #5664c8;
border-radius:8px; background:#5664c8; color:white; font:inherit; cursor:pointer;}
button:hover {background:#4653b4;}
button:focus-visible {outline:3px solid #a9b2f0; outline-offset:2px;}
button:disabled {opacity:.7; cursor:wait;}
[role="status"] {font-size:.85rem; margin-top:.35rem; color:var(--st-text-color);}
textarea {box-sizing:border-box; width:100%; margin-top:.4rem; font:inherit;}
''',
    js='''export default function({parentElement, data}) {
    const button = parentElement.querySelector("button");
    const status = parentElement.querySelector('[role="status"]');
    const fallback = parentElement.querySelector("textarea");
    // User text is passed as data, never interpolated into HTML or JavaScript.
    button.textContent = data.label;
    status.textContent = "";
    fallback.hidden = true;
    fallback.value = data.text;
    let active = true;
    const copy = async () => {
        button.disabled = true;
        try {
            await navigator.clipboard.writeText(data.text);
            if (active) status.textContent = "복사했습니다. K-에듀파인에 붙여 넣으세요.";
        } catch {
            if (active) {
                status.textContent = "브라우저에서 복사를 허용하지 않았습니다. 아래 문구를 선택해 직접 복사해 주세요.";
                fallback.hidden = false;
                fallback.focus();
                fallback.select();
            }
        } finally {
            if (active) button.disabled = false;
        }
    };
    button.addEventListener("click", copy);
    return () => { active = false; button.removeEventListener("click", copy); };
}''',
)


def copy_button(text: str, *, key: str, label="개요 텍스트 복사하기"):
    _copy(data={"text": text, "label": label}, key=key)
