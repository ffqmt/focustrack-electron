window.FocusTrackMockDb = (() => {
  const tickets = [
    {
      id: 3482,
      title: 'Ajustar automação do aceite do chamado',
      team: 'Suporte Interno',
      project: 'Interno',
      status: 'Em andamento',
      timerSeconds: 5263,
      needsApproval: true
    },
    {
      id: 3471,
      title: 'Criar task no Planner ao gerar chamado',
      team: 'Suporte Interno',
      project: 'Planner',
      status: 'Novo',
      timerSeconds: 0,
      needsApproval: true
    },
    {
      id: 3498,
      title: 'Validar relatório semanal de chamados',
      team: 'Operações',
      project: 'Relatórios',
      status: 'Pausado',
      timerSeconds: 1840,
      needsApproval: true
    },
    {
      id: 3520,
      title: 'Ajustar fluxo planner para aceite interno',
      team: 'Suporte Interno',
      project: 'Planner',
      status: 'Reaberto',
      timerSeconds: 2640,
      needsApproval: true
    }
  ];

  function clone(data) {
    return JSON.parse(JSON.stringify(data));
  }

  function searchTickets(query) {
    const normalized = query.trim().toLowerCase().replace('#', '');

    if (!normalized) return [];

    return clone(
      tickets.filter((ticket) => {
        return (
          ticket.title.toLowerCase().includes(normalized) ||
          String(ticket.id).includes(normalized) ||
          ticket.status.toLowerCase().includes(normalized)
        );
      })
    );
  }

  function getTicketById(id) {
    return clone(tickets.find((ticket) => ticket.id === Number(id)) || null);
  }

  function getCurrentActiveTicket() {
    const active = tickets.find((ticket) => ticket.status === 'Em andamento');
    return clone(active || tickets[0] || null);
  }

  function createTicket(title) {
    const nextId = Math.max(...tickets.map((t) => t.id)) + 1;
    const ticket = {
      id: nextId,
      title,
      team: 'Suporte Interno',
      project: 'Geral',
      status: 'Novo',
      timerSeconds: 0,
      needsApproval: true
    };

    tickets.unshift(ticket);
    return clone(ticket);
  }

  function updateTicketStatus(id, status) {
    const index = tickets.findIndex((ticket) => ticket.id === Number(id));
    if (index === -1) return null;
    tickets[index].status = status;
    return clone(tickets[index]);
  }

  function incrementTicketTimer(id) {
    const ticket = tickets.find((t) => t.id === Number(id));
    if (!ticket) return null;
    ticket.timerSeconds += 1;
    return clone(ticket);
  }

  return {
    searchTickets,
    getTicketById,
    getCurrentActiveTicket,
    createTicket,
    updateTicketStatus,
    incrementTicketTimer
  };
})();
