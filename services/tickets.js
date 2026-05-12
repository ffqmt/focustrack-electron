window.FocusTrackTickets = (() => {
  function getConfig() {
    return window.FocusTrackConfigStore?.load?.() || window.FocusTrackConfig || {};
  }

  async function searchTickets(query) {
    const config = getConfig();

    if (config.mode === 'mock' && window.FocusTrackMockDb) {
      return window.FocusTrackMockDb.searchTickets(query);
    }

    const result = await window.focusTrack.searchTickets(query);

    if (!result?.ok) {
      return [];
    }

    return (result.tickets || []).map(ticket => ({
      id: Number(ticket.id),
      title: ticket.title,
      team: ticket.department || ticket.type || 'Geral',
      project: ticket.type || 'Geral',
      status: ticket.status || 'Novo',
      timerSeconds: 0,
      needsApproval: true,
      raw: ticket
    }));
  }

  async function createTicket(title) {
    const config = getConfig();

    if (config.mode === 'mock' && window.FocusTrackMockDb) {
      return window.FocusTrackMockDb.createTicket(title);
    }

    const result = await window.focusTrack.createTicket({
      title,
      status: 'Novo',
      type: 'Chamado',
      origin: 'FocusTrack'
    });

    if (!result?.ok) {
      return null;
    }

    const item = result.ticket;
    const fields = item?.fields || {};

    return {
      id: Number(item.id),
      title: fields.Title || title,
      team: fields.Departamento || 'Geral',
      project: fields.TipodeChamado || 'Chamado',
      status: fields.Status || 'Novo',
      timerSeconds: 0,
      needsApproval: true,
      raw: item
    };
  }

  async function getCurrentActiveTicket() {
    const tickets = await searchTickets('');
    return (
      tickets.find(ticket => ticket.status === 'Em andamento') ||
      tickets[0] ||
      null
    );
  }

  async function getTicketById(id) {
    const tickets = await searchTickets('');
    return tickets.find(ticket => Number(ticket.id) === Number(id)) || null;
  }

  return {
    searchTickets,
    createTicket,
    getCurrentActiveTicket,
    getTicketById
  };
})();
