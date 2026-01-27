# Outlook Sidebar - Feature Backlog

This document captures high-value features for future versions, organized by functional area.

---

## 1) Flagged / Reminders Pane (Bottom Pane Enhancement)

**Goal:** Make the bottom pane earn its 50% space with a comprehensive unified view.

### Unified List
- [ ] Implement tabs or sections for:
  - Flagged Mail
  - Tasks
  - Meetings
  - Reminders firing soon

### Due Grouping
- [ ] Group items by due date:
  - Overdue
  - Today
  - This Week
  - Later
  - No due date

### Sort Controls
- [ ] Add sorting options:
  - Due time
  - Created/received time
  - Importance
  - Sender/organiser

### Quick Actions Per Item Type

**Flag Actions:**
- [ ] Complete
- [ ] Clear
- [ ] Change due date:
  - Today
  - Tomorrow
  - This week
  - Custom date picker

**Task Actions:**
- [ ] Complete
- [ ] Defer
- [ ] Set reminder time

**Meeting Actions:**
- [ ] Join/open
- [ ] Accept/Decline/Tentative (optional)

### Next Reminder Strip
- [ ] Display next 1-3 reminders due to pop
- [ ] Show countdown/time until reminder fires

---

## 2) Email Cards (Enhanced Actionability)

### Attachment Features
- [ ] Attachment pill with file count
- [ ] "Open latest attachment" shortcut
- [ ] Attachment type icons

### Link Detection
- [ ] Smart link detection (Smartsheet URLs, etc.)
- [ ] Show links as clickable
- [ ] Copy button for detected links

### Conversation Indicators
- [ ] Unread count in thread
- [ ] Last reply time
- [ ] Thread depth indicator

### Preview Controls
- [ ] Expand/collapse preview on hover/click
- [ ] Keep list compact by default
- [ ] Smooth expand/collapse animation

### Sender Management
- [ ] Sender trust / VIP markers
- [ ] Pin favourite senders
- [ ] Sender avatars/initials

---

## 3) Actions & Button System

### Custom Button Presets
- [ ] Per-mode button configurations:
  - Email mode
  - Flagged mode
  - Tasks mode

### Additional Built-in Actions
- [ ] Flag/Unflag
- [ ] Complete flag
- [ ] Move to folder (leverage existing "Folders (for Move)" column)
- [ ] Archive
- [ ] Copy subject
- [ ] Copy link
- [ ] Copy sender
- [ ] **Create task from email** (high productivity win)
- [ ] **Create calendar reminder from email**

### Multi-Select & Bulk Actions
- [ ] Multi-select UI (checkboxes)
- [ ] Bulk operations:
  - Mark read/unread
  - Delete
  - Move to folder
  - Flag/unflag
  - Apply category

---

## 4) Filtering & Saved Views

### Saved Filters
- [ ] Predefined views:
  - "Support only"
  - "VIPs"
  - "Overdue follow-ups"
  - "Unread w/ attachments"
- [ ] Custom filter builder
- [ ] Save/load filter presets

### Rules-Lite Exclusions
- [ ] Hide newsletters
- [ ] Ignore subjects containing X
- [ ] Ignore categories Y
- [ ] Regex-based exclusions

### Per-Folder Monitoring
- [ ] Monitor Inbox + chosen folder list
- [ ] Multiple mailbox support
- [ ] Folder selection UI

### Search Bar
- [ ] Instant filter across:
  - Subject
  - Sender
  - Snippet/body preview
- [ ] Search history

---

## 5) Notifications That Are Actually Helpful

### Local Toast Notifications
- [ ] New email from VIPs
- [ ] Flag due soon
- [ ] Flag overdue
- [ ] Reminder firing in X minutes

### Do Not Disturb
- [ ] During meetings (auto-detect)
- [ ] Outside hours (configurable time ranges)
- [ ] Manual toggle

### Escalation
- [ ] "If overdue > 24h, keep pinned + badge count"
- [ ] Configurable escalation rules
- [ ] Visual urgency indicators

---

## 6) Layout & UX Polish

### Window Management
- [ ] Resizable split between Email and Flagged/Reminders (drag handle)
- [ ] Per-monitor docking
- [ ] Remember window position per monitor
- [ ] Always-on-top toggle
- [ ] Auto-hide option

### Display Modes
- [ ] Compact/dense mode
- [ ] "Big card" mode
- [ ] Adjustable font sizes
- [ ] Theme customization

### Keyboard Shortcuts
- [ ] `j/k` - Navigation (next/previous)
- [ ] `Enter` - Open item
- [ ] `r` - Reply
- [ ] `d` - Delete
- [ ] `f` - Flag
- [ ] `c` - Complete
- [ ] `Esc` - Close/deselect
- [ ] `/` - Focus search
- [ ] Customizable shortcuts

---

## 7) Performance & Reliability (COM Realities)

### Event-Driven Architecture
- [ ] Use `ItemAdd`/`ItemChange` events instead of strict polling where possible
- [ ] Incremental updates instead of full list rebuild
- [ ] Debounce rapid changes

### Graceful Degradation
- [ ] "Outlook not running" state with clear messaging
- [ ] Auto-reconnect on Outlook restart
- [ ] Retry logic with exponential backoff

### Caching & Optimization
- [ ] Cache email metadata
- [ ] Incremental updates only for changed items
- [ ] Lazy loading for large lists
- [ ] Virtual scrolling for performance

### Multi-Profile Support
- [ ] Choose store/mailbox
- [ ] Shared mailbox support
- [ ] Multiple account handling

---

## 8) "Nice Later" Power Features

### Stale Follow-Up Detection
- [ ] Detect flagged items with no activity for X days
- [ ] Suggest archiving or re-prioritizing
- [ ] Configurable staleness threshold

### Time-to-Clear Stats
- [ ] Track "Average time flagged → completed"
- [ ] Productivity metrics dashboard
- [ ] Weekly/monthly reports

### Quick Notes on Items
- [ ] Add notes to emails/tasks
- [ ] Store in user property or category convention
- [ ] Search notes

### Integration Hooks
- [ ] Detect Smartsheet/Jira/etc. links in subject/body
- [ ] Show mini badge for detected integrations
- [ ] Quick "open in X" button
- [ ] Custom integration plugins

### Advanced Features
- [ ] Email templates
- [ ] Quick reply snippets
- [ ] Auto-categorization based on rules
- [ ] Machine learning for priority prediction
- [ ] Natural language due date parsing ("tomorrow", "next Friday")

---

## Notes

- This backlog was generated from GPT suggestions based on the current UI
- Features are organized by functional area for easier planning
- Prioritization should be based on user feedback and development effort
- Some features may require significant COM API research
- Consider breaking large features into smaller incremental improvements
