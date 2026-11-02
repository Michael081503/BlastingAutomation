export default function ({ data, parentElement, setTriggerValue }) {
    const list = parentElement.querySelector('.agent-list');
    const status = parentElement.querySelector('.status');
    let agents = data.map(agent => ({ ...agent }));
    let draggedCode = null;

    function move(code, targetIndex, focusDirection = null) {
        const sourceIndex = agents.findIndex(agent => agent.code === code);
        if (sourceIndex < 0 || sourceIndex === targetIndex || targetIndex < 0 || targetIndex >= agents.length) return;
        const [agent] = agents.splice(sourceIndex, 1);
        agents.splice(targetIndex, 0, agent);
        render();
        status.textContent = `${agent.name} moved to position ${targetIndex + 1}. Saving order…`;
        if (focusDirection) {
            const row = [...list.children].find(item => item.dataset.code === code);
            const button = row.querySelector(`[data-direction="${focusDirection}"]`);
            (button.disabled ? row.querySelector('button:not(:disabled)') : button).focus();
        }
        setTriggerValue('order', agents.map(agent => agent.code));
    }

    function render() {
        list.replaceChildren();
        agents.forEach((agent, index) => {
            const row = document.createElement('li');
            row.className = 'agent-row';
            row.draggable = true;
            row.dataset.code = agent.code;
            const handle = document.createElement('span');
            handle.className = 'handle';
            handle.textContent = '⠿';
            handle.setAttribute('aria-hidden', 'true');
            const position = document.createElement('span');
            position.className = 'position';
            position.textContent = String(index + 1);
            const label = document.createElement('span');
            label.className = 'agent-label';
            const name = document.createElement('strong');
            name.textContent = agent.name;
            const code = document.createElement('small');
            code.textContent = agent.code;
            label.append(name, code);
            const actions = document.createElement('span');
            actions.className = 'actions';
            for (const [direction, delta, arrow] of [['up', -1, '↑'], ['down', 1, '↓']]) {
                const button = document.createElement('button');
                button.type = 'button';
                button.dataset.direction = direction;
                button.textContent = arrow;
                button.title = `Move ${agent.name} ${direction}`;
                button.setAttribute('aria-label', button.title);
                button.disabled = index + delta < 0 || index + delta >= agents.length;
                button.onclick = () => move(agent.code, index + delta, direction);
                actions.appendChild(button);
            }
            row.append(handle, position, label, actions);
            row.addEventListener('dragstart', event => {
                if (event.target.closest('button')) {
                    event.preventDefault();
                    return;
                }
                draggedCode = agent.code;
                event.dataTransfer.effectAllowed = 'move';
                event.dataTransfer.setData('text/plain', agent.code);
                row.classList.add('dragging');
            });
            row.addEventListener('dragover', event => {
                if (draggedCode === null || draggedCode === agent.code) return;
                event.preventDefault();
                event.dataTransfer.dropEffect = 'move';
                row.classList.add('drop-target');
            });
            row.addEventListener('dragleave', () => row.classList.remove('drop-target'));
            row.addEventListener('drop', event => {
                event.preventDefault();
                row.classList.remove('drop-target');
                if (draggedCode === null) return;
                const code = draggedCode;
                draggedCode = null;
                move(code, index);
            });
            row.addEventListener('dragend', () => {
                draggedCode = null;
                list.querySelectorAll('.agent-row').forEach(item => item.classList.remove('dragging', 'drop-target'));
            });
            list.appendChild(row);
        });
    }

    render();
}
