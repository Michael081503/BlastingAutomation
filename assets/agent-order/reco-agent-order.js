export default function createAgentOrder({ parentElement, onOrderChange }) {
    const list = parentElement.querySelector('.agent-list');
    const status = parentElement.querySelector('.status');
    const rows = new Map();
    let agents = [];
    let revision = null;
    let saving = false;
    let draggedCode = null;

    function move(code, targetIndex, focusDirection = null) {
        const sourceIndex = agents.findIndex(agent => agent.code === code);
        if (saving || sourceIndex < 0 || sourceIndex === targetIndex || targetIndex < 0 || targetIndex >= agents.length) return;
        const [agent] = agents.splice(sourceIndex, 1);
        agents.splice(targetIndex, 0, agent);
        saving = true;
        render();
        status.textContent = 'Saving order...';
        if (focusDirection) {
            const row = rows.get(code);
            const button = row.querySelector(`[data-direction="${focusDirection}"]`);
            (button.disabled ? row.querySelector('button:not(:disabled)') : button).focus({ preventScroll: true });
        }
        onOrderChange({ order: agents.map(agent => agent.code), revision, event_id: `${Date.now()}-${Math.random()}` });
    }

    function createRow(agent) {
        const row = document.createElement('li');
        row.className = 'agent-row';
        row.dataset.code = agent.code;
        const handle = document.createElement('span');
        handle.className = 'handle';
        handle.textContent = '\u283f';
        handle.setAttribute('aria-hidden', 'true');
        const position = document.createElement('span');
        position.className = 'position';
        const label = document.createElement('span');
        label.className = 'agent-label';
        const name = document.createElement('strong');
        name.textContent = agent.name;
        const code = document.createElement('small');
        code.textContent = agent.code;
        label.append(name, code);
        const actions = document.createElement('span');
        actions.className = 'actions';
        for (const [direction, delta, arrow] of [['up', -1, '\u2191'], ['down', 1, '\u2193']]) {
            const button = document.createElement('button');
            button.type = 'button';
            button.dataset.direction = direction;
            button.textContent = arrow;
            button.title = `Move ${agent.name} ${direction}`;
            button.setAttribute('aria-label', button.title);
            button.onclick = () => move(agent.code, agents.findIndex(item => item.code === agent.code) + delta, direction);
            actions.appendChild(button);
        }
        row.append(handle, position, label, actions);
        row.addEventListener('dragstart', event => {
            if (saving || event.target.closest('button')) {
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
            const code = draggedCode;
            draggedCode = null;
            list.querySelectorAll('.agent-row').forEach(item => item.classList.remove('dragging', 'drop-target'));
            if (code !== null) move(code, agents.findIndex(item => item.code === agent.code));
        });
        row.addEventListener('dragend', () => {
            draggedCode = null;
            list.querySelectorAll('.agent-row').forEach(item => item.classList.remove('dragging', 'drop-target'));
        });
        return row;
    }

    function render() {
        agents.forEach((agent, index) => {
            if (!rows.has(agent.code)) rows.set(agent.code, createRow(agent));
            const row = rows.get(agent.code);
            row.draggable = !saving;
            row.querySelector('.position').textContent = String(index + 1);
            row.querySelector('[data-direction="up"]').disabled = index === 0;
            row.querySelector('[data-direction="down"]').disabled = index === agents.length - 1;
            if (list.children[index] !== row) list.insertBefore(row, list.children[index] || null);
        });
    }

    return function update(data, savedRevision, saveError) {
        const wasSaving = saving;
        agents = data.map(agent => ({ ...agent }));
        revision = savedRevision;
        saving = false;
        render();
        if (saveError) status.textContent = 'Order could not be saved. Try again.';
        else if (wasSaving) status.textContent = 'Order saved.';
        else if (!status.textContent) status.textContent = 'Order saves automatically.';
    };
}
