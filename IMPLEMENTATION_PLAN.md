# Outlook Sidebar - Prioritized Implementation Plan

This document outlines a **sense-checked, phased implementation roadmap** based on the backlog features, current architecture (v1.0.8), and realistic development effort.

---

## 🎯 Guiding Principles

1. **Build on existing strengths** - You already have Split View, Live Settings, and COM integration working
2. **High value, low complexity first** - Quick wins that users will notice immediately
3. **Incremental delivery** - Each phase should be shippable and testable
4. **Respect COM limitations** - Some features require significant API research
5. **User feedback loops** - Ship, test, iterate before moving to complex features

---

## Phase 1: Quick Wins - Email Card Enhancements
**Effort:** Low | **Value:** High | **Target:** v1.1.x

These build directly on your existing 3-line card layout and require minimal architectural changes.

### 1.1 Enhanced Attachment Handling
- [x] Visual attachment indicator (@ symbol) - **DONE in v1.0.8**
- [ ] Attachment count badge (e.g., "@ 3" instead of just "@")
- [ ] Attachment type icons (PDF, Excel, Word, Image)
- [ ] "Open latest attachment" quick action button
- [ ] File size display for attachments

**Why first:** You've already got attachment detection working. This is just UI polish.

### 1.2 Sender VIP/Trust Markers
- [ ] Add a "VIP" star icon next to sender name for flagged contacts
- [ ] Settings: Add/remove VIP senders (simple text list)
- [ ] Visual differentiation (gold star, different color)

**Why first:** Simple string matching, no complex COM calls. High productivity value.

### 1.3 Conversation Thread Indicators
- [ ] Show unread count in thread (e.g., "3 unread in conversation")
- [ ] Last reply timestamp
- [ ] Thread depth indicator (simple icon)

**Why first:** Leverages existing `ConversationTopic` property. Minimal COM overhead.

---

## Phase 2: Actions & Productivity - The Big Wins
**Effort:** Medium | **Value:** Very High | **Target:** v1.2.x

These are the features that will make users say "I can't live without this."

### 2.1 Essential Quick Actions
- [ ] **Flag/Unflag** - Single-click flag toggle
- [ ] **Complete flag** - Mark flagged item as complete
- [ ] **Copy subject** - Clipboard integration
- [ ] **Copy sender** - Quick email address copy
- [ ] **Archive** - Move to Archive folder (configurable)

**Why second:** These are simple COM operations. High frequency use cases.

### 2.2 Game-Changer Actions
- [ ] **Create task from email** - Convert email to Outlook Task
  - Auto-populate subject, due date, body snippet
  - Link back to original email
- [ ] **Create calendar reminder from email** - Quick meeting/reminder creation
  - Parse dates from subject/body if possible
  - Default to "tomorrow" if no date detected

**Why second:** These are the "productivity multiplier" features. Worth the COM research.

### 2.3 Move to Folder Enhancement
- [ ] Implement "Move to Folder" action (you already have "Folders (for Move)" in settings)
- [ ] Quick folder picker (recent folders + favorites)
- [ ] Keyboard shortcut for common folders

**Why second:** Builds on existing folder infrastructure.

---

## Phase 3: Flagged/Reminders Pane - Make It Useful
**Effort:** Medium-High | **Value:** High | **Target:** v1.3.x

You've got the split view. Now make the bottom pane earn its space.

### 3.1 Unified Flagged Items List
- [ ] Fetch all flagged emails (already partially done)
- [ ] Fetch Outlook Tasks
- [ ] Fetch upcoming meetings (next 7 days)
- [ ] Combine into single scrollable list

**Why third:** You've already got the UI container. This is "just" data fetching.

### 3.2 Due Date Grouping
- [ ] Group by: Overdue / Today / Tomorrow / This Week / Later / No Due Date
- [ ] Collapsible sections with counts
- [ ] Visual urgency indicators (red for overdue, amber for today)

**Why third:** Builds on existing `TaskDueDate` filtering from Phase 7.5.

### 3.3 Quick Actions for Flagged Items
- [ ] Complete/Clear flag
- [ ] Defer (Tomorrow/This Week/Custom)
- [ ] Snooze reminder
- [ ] Delete/Archive

**Why third:** Reuses action button framework from email cards.

### 3.4 Next Reminder Strip
- [ ] Top-of-pane strip showing next 1-3 reminders
- [ ] Countdown timer ("Due in 2 hours")
- [ ] Quick complete/snooze buttons

**Why third:** High visibility feature. Builds on existing reminder logic.

---

## Phase 4: Filtering & Search - Taming the Noise
**Effort:** Medium | **Value:** High | **Target:** v1.4.x

### 4.1 Search Bar
- [ ] Add search input at top of email pane
- [ ] Instant filter by subject/sender/snippet
- [ ] Clear button
- [ ] Keyboard shortcut (`/` to focus)

**Why fourth:** You already have filtering logic. This is just a UI wrapper.

### 4.2 Saved Filter Presets
- [ ] Predefined views: "VIPs", "Unread w/ Attachments", "Flagged Only"
- [ ] Quick toggle buttons in header
- [ ] Save current filter as preset

**Why fourth:** Builds on existing "Include read email" and "Show if has Attachment" logic.

### 4.3 Rules-Lite Exclusions
- [ ] Settings: "Ignore subjects containing..." (comma-separated list)
- [ ] "Ignore senders..." (email list)
- [ ] "Ignore categories..." (category picker)

**Why fourth:** Simple string matching. No complex COM calls.

---

## Phase 5: Keyboard Shortcuts - Power User Mode
**Effort:** Low-Medium | **Value:** Medium-High | **Target:** v1.5.x

### 5.1 Core Navigation
- [ ] `j/k` - Next/Previous email
- [ ] `Enter` - Open selected email
- [ ] `Esc` - Deselect/Close
- [ ] `r` - Reply
- [ ] `d` - Delete
- [ ] `f` - Flag/Unflag
- [ ] `c` - Complete flag
- [ ] `/` - Focus search

**Why fifth:** Tkinter keyboard binding is straightforward. High power-user value.

### 5.2 Customizable Shortcuts
- [ ] Settings: Remap shortcuts
- [ ] Conflict detection
- [ ] Reset to defaults

**Why fifth:** Nice-to-have after core shortcuts work.

---

## Phase 6: Layout & UX Polish
**Effort:** Medium | **Value:** Medium | **Target:** v1.6.x

### 6.1 Resizable Split
- [ ] Drag handle between Email and Flagged/Reminders panes
- [ ] Remember split ratio in settings
- [ ] Double-click to reset to 50/50

**Why sixth:** You already have grid-based split. This adds interactivity.

### 6.2 Compact/Dense Mode
- [ ] Toggle between normal and compact card sizes
- [ ] Reduce padding, smaller fonts
- [ ] Show more emails on screen

**Why sixth:** Simple CSS-like adjustments to existing card rendering.

### 6.3 Preview Expand/Collapse
- [ ] Click card to expand body preview
- [ ] Smooth animation
- [ ] Collapse others when expanding new card

**Why sixth:** Nice UX polish. Not critical for functionality.

---

## Phase 7: Performance & Reliability
**Effort:** High | **Value:** High | **Target:** v1.7.x

### 7.1 Event-Driven Updates
- [ ] Research `ItemAdd`/`ItemChange` COM events
- [ ] Replace polling with event listeners where possible
- [ ] Keep polling as fallback

**Why seventh:** Requires significant COM research. High technical risk.

### 7.2 Caching & Incremental Updates
- [ ] Cache email metadata in memory
- [ ] Only update changed items
- [ ] Reduce full list rebuilds

**Why seventh:** Performance optimization. Do this after feature set stabilizes.

### 7.3 Virtual Scrolling
- [ ] Render only visible cards
- [ ] Lazy load off-screen items
- [ ] Handle large inboxes (1000+ emails)

**Why seventh:** Complex Tkinter work. Only needed if performance becomes an issue.

---

## Phase 8: Notifications & DND
**Effort:** Medium | **Value:** Medium-High | **Target:** v1.8.x

### 8.1 Smart Notifications
- [ ] Toast for new VIP emails
- [ ] Toast for overdue flags
- [ ] Toast for upcoming reminders (15 min warning)

**Why eighth:** You already have notification infrastructure. This is just smarter triggers.

### 8.2 Do Not Disturb
- [ ] Auto-DND during meetings (check Outlook calendar)
- [ ] Manual DND toggle in header
- [ ] Scheduled DND (e.g., 6pm-8am)

**Why eighth:** Requires calendar integration. Medium complexity.

---

## Phase 9: "Nice Later" Power Features
**Effort:** High | **Value:** Medium | **Target:** v2.0+

These are cool but not essential. Save for when the core product is rock-solid.

### 9.1 Link Detection & Smart Actions
- [ ] Detect Smartsheet/Jira/GitHub URLs
- [ ] Show badge + "Open in X" button
- [ ] Copy link button

### 9.2 Stale Follow-Up Detection
- [ ] Highlight flagged items with no activity for 7+ days
- [ ] Suggest archiving or re-prioritizing

### 9.3 Productivity Stats
- [ ] Track "Average time flagged → completed"
- [ ] Weekly summary dashboard
- [ ] Export to CSV

### 9.4 Quick Notes
- [ ] Add notes to emails (stored in custom property)
- [ ] Search notes
- [ ] Show note indicator on cards

---

## 🚫 Features to Avoid (For Now)

These are interesting but have significant technical or UX challenges:

- **Multi-select & bulk actions** - Complex Tkinter state management. Wait until single-item actions are solid.
- **Multi-profile/shared mailbox support** - Significant COM complexity. Niche use case.
- **Machine learning priority prediction** - Overkill for v1.x. Focus on manual controls first.
- **Email templates & quick replies** - Outlook already has this. Don't reinvent the wheel.

---

## 📊 Summary Roadmap

| Phase | Focus Area | Effort | Value | Target Version |
|-------|-----------|--------|-------|----------------|
| 1 | Email Card Enhancements | Low | High | v1.1.x |
| 2 | Actions & Productivity | Medium | Very High | v1.2.x |
| 3 | Flagged/Reminders Pane | Medium-High | High | v1.3.x |
| 4 | Filtering & Search | Medium | High | v1.4.x |
| 5 | Keyboard Shortcuts | Low-Medium | Medium-High | v1.5.x |
| 6 | Layout & UX Polish | Medium | Medium | v1.6.x |
| 7 | Performance & Reliability | High | High | v1.7.x |
| 8 | Notifications & DND | Medium | Medium-High | v1.8.x |
| 9 | Power Features | High | Medium | v2.0+ |

---

## 🎯 Recommended Next Steps

**Start with Phase 1.1-1.2** (Attachment enhancements + VIP markers):
1. Add attachment count to existing @ indicator
2. Implement VIP sender list in settings
3. Show gold star for VIP senders
4. Ship as v1.1.0

**Then move to Phase 2.1** (Essential quick actions):
1. Add Flag/Unflag button
2. Add Copy Subject/Sender buttons
3. Ship as v1.2.0

**User feedback checkpoint:** See which actions get used most before building Phase 2.2.

---

## 💡 Key Insights

1. **Your Split View is ahead of schedule** - You've already built the container for Phase 3. Focus on filling it with useful data.

2. **COM is your bottleneck** - Features requiring new COM API calls (Tasks, Meetings, Events) will take longer. Do the research upfront.

3. **Settings infrastructure is solid** - Your Live Settings panel makes adding new toggles/filters very easy. Leverage this.

4. **Don't over-engineer** - Ship incremental improvements. Get user feedback. Iterate.

5. **Phase 2.2 is the killer feature** - "Create task from email" will be a game-changer. Worth the investment.
